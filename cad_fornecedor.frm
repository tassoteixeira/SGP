VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form cadastro_fornecedor 
   Caption         =   "Cadastro de Fornecedores"
   ClientHeight    =   6795
   ClientLeft      =   1170
   ClientTop       =   1065
   ClientWidth     =   7035
   Icon            =   "cad_fornecedor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "cad_fornecedor.frx":030A
   ScaleHeight     =   6795
   ScaleWidth      =   7035
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   120
      Picture         =   "cad_fornecedor.frx":0750
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Cria um novo registro."
      Top             =   5820
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   1020
      Picture         =   "cad_fornecedor.frx":1DE2
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Altera o registro atual."
      Top             =   5820
      Width           =   795
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   1920
      Picture         =   "cad_fornecedor.frx":32DC
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Exclui o registro atual."
      Top             =   5820
      Width           =   795
   End
   Begin VB.CommandButton cmd_pesquisa 
      Caption         =   "&Pesquisa"
      Height          =   855
      Left            =   2820
      Picture         =   "cad_fornecedor.frx":496E
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Pesquisa um registro específico."
      Top             =   5820
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3720
      Picture         =   "cad_fornecedor.frx":5DE0
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   5820
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Enabled         =   0   'False
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   6795
      Begin VB.TextBox txt_conta_contabil 
         Height          =   285
         Left            =   1740
         MaxLength       =   8
         TabIndex        =   31
         Top             =   5280
         Width           =   915
      End
      Begin VB.TextBox txt_inscricao_estadual 
         Height          =   285
         Left            =   1740
         MaxLength       =   13
         TabIndex        =   26
         Top             =   4560
         Width           =   2355
      End
      Begin VB.TextBox txt_vendedor 
         Height          =   285
         Left            =   1740
         MaxLength       =   40
         TabIndex        =   22
         Top             =   3840
         Width           =   4935
      End
      Begin VB.TextBox txt_fax 
         Height          =   285
         Left            =   1740
         MaxLength       =   15
         TabIndex        =   20
         Top             =   3480
         Width           =   1455
      End
      Begin VB.TextBox txt_fone2 
         Height          =   285
         Left            =   1740
         MaxLength       =   15
         TabIndex        =   18
         Top             =   3120
         Width           =   1455
      End
      Begin VB.TextBox txt_fone 
         Height          =   285
         Left            =   1740
         MaxLength       =   15
         TabIndex        =   16
         Top             =   2760
         Width           =   1455
      End
      Begin VB.TextBox txt_cidade 
         Height          =   285
         Left            =   1740
         MaxLength       =   30
         TabIndex        =   10
         Top             =   1680
         Width           =   3795
      End
      Begin VB.TextBox txt_bairro 
         Height          =   285
         Left            =   1740
         MaxLength       =   30
         TabIndex        =   8
         Top             =   1320
         Width           =   3795
      End
      Begin VB.TextBox txt_endereco 
         Height          =   285
         Left            =   1740
         MaxLength       =   40
         TabIndex        =   6
         Top             =   960
         Width           =   4935
      End
      Begin VB.ComboBox cbo_conta 
         Height          =   300
         ItemData        =   "cad_fornecedor.frx":7472
         Left            =   2160
         List            =   "cad_fornecedor.frx":7474
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   4920
         Width           =   4515
      End
      Begin VB.TextBox txt_uf 
         Height          =   285
         Left            =   1740
         MaxLength       =   2
         TabIndex        =   12
         Top             =   2040
         Width           =   435
      End
      Begin VB.TextBox txt_codigo 
         Height          =   285
         Left            =   1740
         MaxLength       =   4
         TabIndex        =   2
         Top             =   240
         Width           =   435
      End
      Begin VB.TextBox txt_nome 
         Height          =   285
         Left            =   1740
         MaxLength       =   40
         TabIndex        =   4
         Top             =   600
         Width           =   4935
      End
      Begin MSMask.MaskEdBox msk_cgc 
         Height          =   300
         Left            =   1740
         TabIndex        =   24
         Top             =   4200
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
         Left            =   1740
         TabIndex        =   14
         Top             =   2400
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
      Begin VB.Label Label14 
         Caption         =   "&N. da Conta Contábil"
         Height          =   315
         Left            =   120
         TabIndex        =   30
         Top             =   5280
         Width           =   1575
      End
      Begin VB.Label Label13 
         Caption         =   "&Inscrição Estadual"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   4560
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "C.&G.C."
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   4200
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "&Vendedor"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   3840
         Width           =   1575
      End
      Begin VB.Label Label12 
         Caption         =   "&Fax"
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Label Label11 
         Caption         =   "&Telefone 2"
         Height          =   315
         Left            =   120
         TabIndex        =   17
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "&Telefone 1"
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Label10 
         Caption         =   "Ci&dade"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "&Bairro"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "&Endereço"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lbl_conta 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1740
         TabIndex        =   28
         Top             =   4920
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Sigl&a do Estado"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Código do Fo&rnecedor"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Co&nta"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   27
         Top             =   4920
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "No&me"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "C.E.&P."
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Top             =   2400
         Width           =   1575
      End
   End
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   4740
      TabIndex        =   39
      Top             =   5700
      Width           =   2175
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "cad_fornecedor.frx":7476
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "cad_fornecedor.frx":89F8
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "cad_fornecedor.frx":9E6A
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "cad_fornecedor.frx":B364
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   6120
      Picture         =   "cad_fornecedor.frx":C85E
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Cancela o registro atual."
      Top             =   5820
      Width           =   795
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   5220
      Picture         =   "cad_fornecedor.frx":DD58
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Confirma o registro atual."
      Top             =   5820
      Width           =   795
   End
End
Attribute VB_Name = "cadastro_fornecedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lFlagCadastro As Integer
Dim lOpcao As Integer
Dim lCodigo As Integer
Private Conta As New cContas
Private Fornecedor As New cFornecedor
Private Sub AtivaBotoes()
    cmd_novo.Enabled = True
    cmd_excluir.Enabled = True
    cmd_pesquisa.Enabled = True
    cmd_alterar.Enabled = True
    cmd_sair.Enabled = True
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
    frm_move.Visible = True
    txt_codigo.Enabled = True
End Sub
Private Sub AtualTabe()
    lCodigo = Val(txt_codigo.Text)
    Fornecedor.Empresa = g_empresa
    Fornecedor.Codigo = Val(txt_codigo.Text)
    Fornecedor.Nome = txt_nome.Text
    Fornecedor.Endereco = txt_endereco.Text
    Fornecedor.Bairro = txt_bairro.Text
    Fornecedor.Cidade = txt_cidade.Text
    Fornecedor.UF = txt_uf.Text
    Fornecedor.CEP = msk_cep
    Fornecedor.Telefone = fDesmascaraTelefone(txt_fone.Text)
    Fornecedor.Telefone2 = fDesmascaraTelefone(txt_fone2.Text)
    Fornecedor.Fax = fDesmascaraTelefone(txt_fax.Text)
    Fornecedor.Vendedor = txt_vendedor.Text
    Fornecedor.CGC = msk_cgc
    Fornecedor.InscricaoEstadual = txt_inscricao_estadual.Text
    Fornecedor.CodigoConta = cbo_conta.ItemData(cbo_conta.ListIndex)
    Fornecedor.ContaContabil = txt_conta_contabil.Text
End Sub
Private Sub AtualTela()
    Dim i As Integer
    lCodigo = Fornecedor.Codigo
    txt_codigo.Text = Fornecedor.Codigo
    txt_nome.Text = Fornecedor.Nome
    txt_endereco.Text = Fornecedor.Endereco
    txt_bairro.Text = Fornecedor.Bairro
    txt_cidade.Text = Fornecedor.Cidade
    txt_uf.Text = Fornecedor.UF
    msk_cep = Fornecedor.CEP
    txt_fone.Text = fMascaraTelefone(Fornecedor.Telefone)
    txt_fone2.Text = fMascaraTelefone(Fornecedor.Telefone2)
    txt_fax.Text = fMascaraTelefone(Fornecedor.Fax)
    txt_vendedor.Text = Fornecedor.Vendedor
    msk_cgc = Fornecedor.CGC
    txt_inscricao_estadual.Text = Fornecedor.InscricaoEstadual
    lbl_conta.Caption = Fornecedor.CodigoConta
    txt_conta_contabil.Text = Fornecedor.ContaContabil
    
    If Conta.LocalizarCodigo(Fornecedor.CodigoConta) Then
        For i = 0 To cbo_conta.ListCount - 1
            cbo_conta.ListIndex = i
            If cbo_conta.ItemData(i) = Conta.Codigo Then
                Exit For
            End If
        Next
    Else
        cbo_conta.ListIndex = -1
    End If
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
    Set Conta = Nothing
    Set Fornecedor = Nothing
    frm_cadastro.Show
End Sub
Private Sub Inclui()
    lOpcao = 1
    DesativaBotoes
    cmd_novo.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    txt_codigo.Text = 1
    If Fornecedor.LocalizarUltimo(g_empresa) Then
        txt_codigo.Text = Fornecedor.Codigo + 1
    End If
End Sub
Private Sub cbo_conta_Click()
    If cbo_conta.ListIndex <> -1 Then
        lbl_conta = Format(cbo_conta.ItemData(cbo_conta.ListIndex), "00")
    Else
        lbl_conta = "**"
    End If
End Sub
Private Sub cbo_conta_GotFocus()
    SendMessageLong cbo_conta.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_conta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_conta_contabil.SetFocus
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
    If Fornecedor.LocalizarAnterior Then
        AtualTela
    Else
        MsgBox "Início de Arquivo.", vbInformation, "Atenção!"
        cmd_proximo.SetFocus
    End If
End Sub
Private Sub cmd_cancelar_Click()
    LimpaTela
    lOpcao = 0
    If Fornecedor.LocalizarCodigo(g_empresa, lCodigo) Then
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
    txt_nome.Text = ""
    txt_endereco.Text = ""
    txt_bairro.Text = ""
    txt_cidade.Text = ""
    txt_uf.Text = ""
    msk_cep = ""
    txt_fone.Text = ""
    txt_fone2.Text = ""
    txt_fax.Text = ""
    txt_vendedor.Text = ""
    msk_cgc = ""
    txt_inscricao_estadual.Text = ""
    lbl_conta.Caption = ""
    cbo_conta.ListIndex = -1
    txt_conta_contabil.Text = ""
End Sub
Private Sub cmd_excluir_Click()
    If txt_codigo.Text <> "" Then
        If (MsgBox("Deseja Realmente Excluir Este Registro?", 4 + 32 + 256, "Exclusão de Registro!")) = 6 Then
            If Fornecedor.Excluir(g_empresa, txt_codigo.Text) Then
                LimpaTela
                If Fornecedor.LocalizarUltimo(g_empresa) Then
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
            If Fornecedor.Incluir Then
                lCodigo = txt_codigo.Text
            Else
                MsgBox "Não foi possível incluir este registro!", vbInformation, "Erro de Verificação!"
            End If
        ElseIf lOpcao = 2 Then
            AtualTabe
            If Not Fornecedor.Alterar(g_empresa, lCodigo) Then
                MsgBox "Não foi possível alterar este registro!", vbInformation, "Erro de Verificação!"
            End If
        End If
        lOpcao = 0
        Call Fornecedor.LocalizarCodigo(g_empresa, lCodigo)
        AtualTela
        cmd_novo.SetFocus
    End If
    Exit Sub
FileError:
    MsgBox Error
    'ErroArquivo tbl_fornecedor.Name, "Fornecedoro"
    Exit Sub
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If txt_codigo = "" Then
        MsgBox "Informe o código do fornecedor.", vbInformation, "Atenção!"
        txt_codigo.SetFocus
    ElseIf txt_nome = "" Then
        MsgBox "Informe o nome do fornecedor.", vbInformation, "Atenção!"
        txt_nome.SetFocus
    ElseIf txt_endereco = "" Then
        MsgBox "Informe o endereço.", vbInformation, "Atenção!"
        txt_endereco.SetFocus
    ElseIf txt_bairro = "" Then
        MsgBox "Informe o bairro.", vbInformation, "Atenção!"
        txt_bairro.SetFocus
    ElseIf txt_cidade = "" Then
        MsgBox "Informe a cidade.", vbInformation, "Atenção!"
        txt_cidade.SetFocus
    ElseIf txt_uf = "" Then
        MsgBox "Informe a sígla do estado.", vbInformation, "Atenção!"
        txt_uf.SetFocus
    ElseIf msk_cep = "" Then
        MsgBox "Informe o cep.", vbInformation, "Atenção!"
        msk_cep.SetFocus
    ElseIf cbo_conta.ListIndex = -1 Then
        MsgBox "Informe a conta.", vbInformation, "Atenção!"
        cbo_conta.SetFocus
    ElseIf txt_conta_contabil = "" Then
        MsgBox "Informe a conta contábil.", vbInformation, "Atenção!"
        txt_conta_contabil.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_pesquisa_Click()
    consulta_fornecedor.Show 1
    If Len(g_string) > 0 Then
        lCodigo = RetiraGString(1)
        If Fornecedor.LocalizarCodigo(g_empresa, lCodigo) Then
            AtualTela
        End If
    End If
End Sub
Private Sub cmd_primeiro_Click()
    If Fornecedor.LocalizarPrimeiro Then
        AtualTela
        cmd_proximo.SetFocus
    Else
        MsgBox "Não há registro.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    If Fornecedor.LocalizarProximo Then
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
    If Fornecedor.LocalizarUltimo(g_empresa) Then
        AtualTela
        cmd_anterior.SetFocus
    Else
        MsgBox "Não há registro.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub Form_Activate()
    If lFlagCadastro = 0 Then
        DesativaBotoes
        If Fornecedor.LocalizarUltimo(g_empresa) Then
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
    PreencheCboConta
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub PreencheCboConta()
Dim rsConta As adodb.Recordset
Dim xString As String
    Set rsConta = New adodb.Recordset
    xString = "SELECT Codigo, Nome FROM Contas ORDER BY Nome"
    Set rsConta = Conectar.RsConexao(xString)
    cbo_conta.Clear
    With rsConta
        If .RecordCount > 0 Then
            Do Until .EOF
                cbo_conta.AddItem rsConta("Nome").Value
                cbo_conta.ItemData(cbo_conta.NewIndex) = rsConta("Codigo").Value
                .MoveNext
            Loop
        End If
    End With
    rsConta.Close
    Set rsConta = Nothing
End Sub
Private Sub msk_cgc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_inscricao_estadual.SetFocus
    End If
End Sub
Private Sub txt_bairro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_cidade.SetFocus
    End If
End Sub
Private Sub msk_cep_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_fone.SetFocus
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
        txt_nome.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_conta_contabil_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_endereco_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_bairro.SetFocus
    End If
End Sub
Private Sub txt_fax_GotFocus()
    txt_fax = fDesmascaraTelefone(txt_fax)
End Sub
Private Sub txt_fax_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_vendedor.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_fax_LostFocus()
    txt_fax = fMascaraTelefone(txt_fax)
End Sub
Private Sub txt_fone_GotFocus()
    txt_fone.Text = fDesmascaraTelefone(txt_fone.Text)
End Sub
Private Sub txt_fone_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_fone2.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_fone_LostFocus()
    txt_fone.Text = fMascaraTelefone(txt_fone.Text)
End Sub
Private Sub txt_fone2_GotFocus()
    txt_fone2.Text = fDesmascaraTelefone(txt_fone2.Text)
End Sub
Private Sub txt_fone2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_fax.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_fone2_LostFocus()
    txt_fone2.Text = fMascaraTelefone(txt_fone2.Text)
End Sub
Private Sub txt_inscricao_estadual_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_conta.SetFocus
    End If
End Sub
Private Sub txt_nome_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_endereco.SetFocus
    End If
End Sub
Private Sub txt_uf_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_cep.SetFocus
    End If
End Sub
Private Sub txt_vendedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_cgc.SetFocus
    End If
End Sub
