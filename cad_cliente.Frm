VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form cadastro_cliente 
   Caption         =   "Cadastro de Clientes"
   ClientHeight    =   7545
   ClientLeft      =   315
   ClientTop       =   1395
   ClientWidth     =   7035
   Icon            =   "cad_cliente.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   Picture         =   "cad_cliente.frx":030A
   ScaleHeight     =   7545
   ScaleWidth      =   7035
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   120
      Picture         =   "cad_cliente.frx":0750
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "Cria um novo registro."
      Top             =   6600
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   1020
      Picture         =   "cad_cliente.frx":1DE2
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Altera o registro atual."
      Top             =   6600
      Width           =   795
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   1920
      Picture         =   "cad_cliente.frx":32DC
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Exclui o registro atual."
      Top             =   6600
      Width           =   795
   End
   Begin VB.CommandButton cmd_pesquisa 
      Caption         =   "&Pesquisa"
      Height          =   855
      Left            =   2820
      Picture         =   "cad_cliente.frx":496E
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "Pesquisa um registro específico."
      Top             =   6600
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3720
      Picture         =   "cad_cliente.frx":5DE0
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   6600
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Enabled         =   0   'False
      Height          =   6435
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   6795
      Begin VB.TextBox txtInscricaoEstadual 
         Height          =   285
         Left            =   1740
         MaxLength       =   12
         TabIndex        =   27
         Top             =   4200
         Width           =   1635
      End
      Begin VB.TextBox txtFormaPagamento 
         Height          =   285
         Left            =   1740
         MaxLength       =   40
         TabIndex        =   38
         Top             =   6000
         Width           =   4935
      End
      Begin VB.TextBox txtEmail 
         Height          =   285
         Left            =   1740
         MaxLength       =   40
         TabIndex        =   36
         Top             =   5640
         Width           =   4935
      End
      Begin VB.TextBox txt_telefone 
         Height          =   285
         Left            =   1740
         MaxLength       =   15
         TabIndex        =   16
         Top             =   2760
         Width           =   1455
      End
      Begin VB.CheckBox chk_imprime_duplicata 
         Caption         =   "Imp. Duplicata"
         Height          =   255
         Left            =   5340
         TabIndex        =   25
         Top             =   3840
         Width           =   1335
      End
      Begin VB.CheckBox chk_inativo 
         Caption         =   "Cliente Inativo"
         Height          =   255
         Left            =   5340
         TabIndex        =   28
         Top             =   4200
         Width           =   1335
      End
      Begin VB.TextBox txt_orgao_emissor 
         Height          =   285
         Left            =   5460
         MaxLength       =   10
         TabIndex        =   22
         Top             =   3480
         Width           =   1215
      End
      Begin VB.ComboBox cbo_empresa 
         Height          =   315
         ItemData        =   "cad_cliente.frx":7472
         Left            =   1740
         List            =   "cad_cliente.frx":7474
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   5280
         Width           =   4935
      End
      Begin VB.ComboBox cbo_convenio 
         Height          =   300
         ItemData        =   "cad_cliente.frx":7476
         Left            =   1740
         List            =   "cad_cliente.frx":7478
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   4920
         Width           =   4935
      End
      Begin VB.TextBox txt_identidade 
         Height          =   285
         Left            =   1740
         MaxLength       =   14
         TabIndex        =   20
         Top             =   3480
         Width           =   1635
      End
      Begin VB.TextBox txt_cidade 
         Height          =   285
         Left            =   1740
         MaxLength       =   20
         TabIndex        =   10
         Top             =   1680
         Width           =   2775
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
      Begin VB.ComboBox cbo_vencimento 
         Height          =   315
         ItemData        =   "cad_cliente.frx":747A
         Left            =   1740
         List            =   "cad_cliente.frx":747C
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   4560
         Width           =   4935
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
         MaxLength       =   6
         TabIndex        =   2
         Top             =   240
         Width           =   795
      End
      Begin VB.TextBox txt_razao_social 
         Height          =   285
         Left            =   1740
         MaxLength       =   40
         TabIndex        =   4
         Top             =   600
         Width           =   4935
      End
      Begin MSMask.MaskEdBox msk_cpf 
         Height          =   300
         Left            =   1740
         TabIndex        =   18
         Top             =   3120
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   14
         Format          =   "___.___.___-__"
         Mask            =   "###.###.###-##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_cgc 
         Height          =   300
         Left            =   1740
         TabIndex        =   24
         Top             =   3840
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
      Begin VB.Label Label15 
         Caption         =   "Forma de Pagamento"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   6000
         Width           =   1575
      End
      Begin VB.Label Label14 
         Caption         =   "Email"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   5640
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Órgão Emissor"
         Height          =   315
         Left            =   4320
         TabIndex        =   21
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Empresa"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   33
         Top             =   5280
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Convênio"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   31
         Top             =   4920
         Width           =   1575
      End
      Begin VB.Label Label13 
         Caption         =   "Inscrição Estadual"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   4200
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "C.G.C."
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   3870
         Width           =   1575
      End
      Begin VB.Label Label12 
         Caption         =   "Identidade"
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Label Label11 
         Caption         =   "C.P.F."
         Height          =   315
         Left            =   120
         TabIndex        =   17
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Telefone "
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Label10 
         Caption         =   "Cidade"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "Bairro"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Endereço"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Estado"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Código do cliente"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo de Vencimento"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   29
         Top             =   4560
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Razão social"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "C.E.P."
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
      TabIndex        =   46
      Top             =   6480
      Width           =   2175
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "cad_cliente.frx":747E
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "cad_cliente.frx":8978
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "cad_cliente.frx":9E72
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "cad_cliente.frx":B2E4
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   6120
      Picture         =   "cad_cliente.frx":C866
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "Cancela o registro atual."
      Top             =   6600
      Width           =   795
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   5220
      Picture         =   "cad_cliente.frx":DD60
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Confirma o registro atual."
      Top             =   6600
      Width           =   795
   End
End
Attribute VB_Name = "cadastro_cliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lFlagCadastro As Integer
Dim lOpcao As Integer
Dim lCodigo As Long
Private Cliente As New cCliente
Private Convenio As New cConvenio
Private Empresa As New cEmpresa
Private Vencimento As New cVencimento
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
        Cliente.Codigo = CLng(txt_codigo.Text)
    End If
    Cliente.Codigo = CLng(txt_codigo.Text)
    Cliente.RazaoSocial = txt_razao_social.Text
    Cliente.Endereco = txt_endereco.Text
    Cliente.Bairro = txt_bairro.Text
    Cliente.Cidade = txt_cidade.Text
    Cliente.UF = txt_uf.Text
    Cliente.CEP = msk_cep.Text
    Cliente.Telefone = fDesmascaraTelefone(txt_telefone.Text)
    Cliente.CPF = msk_cpf.Text
    Cliente.CGC = msk_cgc.Text
    Cliente.Identidade = txt_identidade.Text
    Cliente.OrgaoEmissor = txt_orgao_emissor.Text
    Cliente.InscricaoEstadual = txtInscricaoEstadual.Text
    Cliente.CodigoVencimento = cbo_vencimento.ItemData(cbo_vencimento.ListIndex)
    Cliente.CodigoConvenio = cbo_convenio.ItemData(cbo_convenio.ListIndex)
    Cliente.Empresa = cbo_empresa.ItemData(cbo_empresa.ListIndex)
    Cliente.Inativo = chk_inativo.Value
    Cliente.ImprimeDuplicata = chk_imprime_duplicata.Value
    Cliente.Email = txtEmail.Text
    Cliente.FormaPagamento = txtFormaPagamento.Text
End Sub
Private Sub AtualTela()
Dim i As Integer
    lCodigo = Cliente.Codigo
    txt_codigo.Text = Cliente.Codigo
    txt_razao_social.Text = Cliente.RazaoSocial
    txt_endereco.Text = Cliente.Endereco
    txt_bairro.Text = Cliente.Bairro
    txt_cidade.Text = Cliente.Cidade
    txt_uf.Text = Cliente.UF
    msk_cep.Text = Cliente.CEP
    txt_telefone.Text = fMascaraTelefone(Cliente.Telefone)
    msk_cpf.Text = Cliente.CPF
    txt_identidade.Text = Cliente.Identidade
    txt_orgao_emissor.Text = Cliente.OrgaoEmissor
    msk_cgc.Text = Cliente.CGC
    txtInscricaoEstadual.Text = Cliente.InscricaoEstadual
       
    For i = 0 To cbo_vencimento.ListCount - 1
        cbo_vencimento.ListIndex = i
        If cbo_vencimento.ItemData(i) = Cliente.CodigoVencimento Then
            Exit For
        End If
    Next
    
    'Else
    '    cbo_vencimento.ListIndex = -1
    'End If
    
    If Convenio.LocalizarCodigo(Cliente.CodigoConvenio) Then
        For i = 0 To cbo_convenio.ListCount - 1
            cbo_convenio.ListIndex = i
            If cbo_convenio.ItemData(i) = Convenio.Codigo Then
                Exit For
            End If
        Next
    Else
        cbo_convenio.ListIndex = -1
    End If
    
    If Empresa.LocalizarCodigo(Cliente.Empresa) Then
        For i = 0 To cbo_empresa.ListCount - 1
            cbo_empresa.ListIndex = i
            If cbo_empresa.ItemData(i) = Empresa.Codigo Then
                Exit For
            End If
        Next
    Else
        cbo_empresa.ListIndex = -1
    End If
    If Cliente.Inativo Then
        chk_inativo.Value = 1
    Else
        chk_inativo.Value = 0
    End If
    If Cliente.ImprimeDuplicata Then
        chk_imprime_duplicata.Value = 1
    Else
        chk_imprime_duplicata.Value = 0
    End If
    txtEmail.Text = Cliente.Email
    txtFormaPagamento.Text = Cliente.FormaPagamento
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
    'tbl_convenio.Close
    'tbl_empresa.Close
    'tbl_vencimento.Close
    Set Cliente = Nothing
    Set Convenio = Nothing
    Set Empresa = Nothing
    Set Vencimento = Nothing
    frm_cadastro.Show
End Sub
Private Sub Inclui()
    lOpcao = 1
    DesativaBotoes
    cmd_novo.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    txt_codigo.Text = 1
    txt_codigo.Text = Cliente.ProximoCodigo
End Sub
'Private Sub cbo_convenio_Click()
'    If cbo_convenio.ListIndex <> -1 Then
'        txt_convenio = Format(cbo_convenio.ItemData(cbo_convenio.ListIndex), "00")
'    Else
'        txt_convenio = "**"
'    End If
'End Sub
Private Sub cbo_convenio_GotFocus()
    SendMessageLong cbo_convenio.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_convenio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_empresa.SetFocus
    End If
End Sub
'Private Sub cbo_empresa_Click()
'    If cbo_empresa.ListIndex <> -1 Then
'        txt_empresa = Format(cbo_empresa.ItemData(cbo_empresa.ListIndex), "00")
'    Else
'        txt_empresa = "**"
'    End If
'End Sub
Private Sub cbo_empresa_GotFocus()
    SendMessageLong cbo_empresa.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_empresa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtEmail.SetFocus
    End If
End Sub
'Private Sub cbo_vencimento_Click()
'    If cbo_vencimento.ListIndex <> -1 Then
'        txt_vencimento = Format(cbo_vencimento.ItemData(cbo_vencimento.ListIndex), "00")
'    Else
'        txt_vencimento = "**"
'    End If
'End Sub
Private Sub cbo_vencimento_GotFocus()
    SendMessageLong cbo_vencimento.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_vencimento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_convenio.SetFocus
    End If
End Sub
Private Sub chk_imprime_duplicata_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtInscricaoEstadual.SetFocus
    End If
End Sub
Private Sub chk_inativo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_vencimento.SetFocus
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
    txt_razao_social.SetFocus
End Sub
Private Sub cmd_anterior_Click()
    If Cliente.LocalizarAnterior Then
        AtualTela
    Else
        MsgBox "Início de Arquivo.", vbInformation, "Atenção!"
        cmd_proximo.SetFocus
    End If
End Sub
Private Sub cmd_cancelar_Click()
    LimpaTela
    lOpcao = 0
    If Cliente.LocalizarCodigo(lCodigo) Then
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
    txt_razao_social.Text = ""
    txt_endereco.Text = ""
    txt_bairro.Text = ""
    txt_cidade.Text = ""
    txt_uf.Text = ""
    msk_cep.Text = "__.___-___"
    txt_telefone.Text = ""
    msk_cpf.Text = "___.___.___-__"
    txt_identidade.Text = ""
    txt_orgao_emissor.Text = ""
    msk_cgc.Text = "__.___.___/____-__"
    chk_imprime_duplicata.Value = 1
    chk_inativo.Value = 0
    txtInscricaoEstadual.Text = ""
    cbo_vencimento.ListIndex = -1
    cbo_convenio.ListIndex = -1
    cbo_empresa.ListIndex = -1
    txtEmail.Text = ""
    txtFormaPagamento.Text = ""
End Sub
Private Sub cmd_excluir_Click()
    If Val(txt_codigo.Text) > 0 Then
        If (MsgBox("Deseja Realmente Excluir Este Registro?", 4 + 32 + 256, "Exclusão de Registro!")) = 6 Then
            If Cliente.Excluir(txt_codigo.Text) Then
                LimpaTela
                If Cliente.LocalizarUltimo Then
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
    txt_razao_social.SetFocus
End Sub
Private Sub cmd_ok_Click()
    On Error GoTo FileError
    If ValidaCampos Then
        AtivaBotoes
        If lOpcao = 1 Then
            AtualTabe
            If Cliente.Incluir Then
                lCodigo = Val(txt_codigo.Text)
            Else
                MsgBox "Não foi possível incluir este registro!", vbInformation, "Erro de Verificação!"
            End If
        ElseIf lOpcao = 2 Then
            AtualTabe
            If Not Cliente.Alterar(lCodigo) Then
                MsgBox "Não foi possível alterar este registro!", vbInformation, "Erro de Verificação!"
            End If
        End If
        lOpcao = 0
        Call Cliente.LocalizarCodigo(lCodigo)
        AtualTela
        cmd_novo.SetFocus
    End If
    Exit Sub
FileError:
    MsgBox Error
    'ErroArquivo tbl_cliente.Name, "Clienteo"
    Exit Sub
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If Not Val(txt_codigo.Text) > 0 Then
        MsgBox "Informe o código do cliente.", vbInformation, "Atenção!"
        txt_codigo.SetFocus
    ElseIf txt_razao_social.Text = "" Then
        MsgBox "Informe a razão social do cliente.", vbInformation, "Atenção!"
        txt_razao_social.SetFocus
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
    ElseIf Val(msk_cep.Text) < 1 Then
        MsgBox "Informe o cep.", vbInformation, "Atenção!"
        msk_cep.SetFocus
    ElseIf txt_telefone.Text = "" Then
        MsgBox "Informe o telefone.", vbInformation, "Atenção!"
        txt_telefone.SetFocus
    ElseIf cbo_vencimento.ListIndex = -1 Then
        MsgBox "Informe o vencimento.", vbInformation, "Atenção!"
        cbo_vencimento.SetFocus
    ElseIf cbo_convenio.ListIndex = -1 Then
        MsgBox "Informe o convênio.", vbInformation, "Atenção!"
        cbo_convenio.SetFocus
    ElseIf cbo_empresa.ListIndex = -1 Then
        MsgBox "Informe a empresa.", vbInformation, "Atenção!"
        cbo_empresa.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_pesquisa_Click()
    consulta_cliente.Show 1
    If Len(g_string) > 0 Then
        lCodigo = RetiraGString(1)
        Call Cliente.LocalizarCodigo(lCodigo)
        AtualTela
    End If
End Sub
Private Sub cmd_primeiro_Click()
    If Cliente.LocalizarPrimeiro Then
        AtualTela
        cmd_proximo.SetFocus
    Else
        MsgBox "Não há registro.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    If Cliente.LocalizarProximo Then
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
    If Cliente.LocalizarUltimo Then
        AtualTela
        cmd_anterior.SetFocus
    Else
        MsgBox "Não há registro.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub Form_Activate()
    If lFlagCadastro = 0 Then
        DesativaBotoes
        If Cliente.LocalizarUltimo Then
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
    PreencheCboVencimento
    PreencheCboConvenio
    PreencheCboEmpresa
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub PreencheCboVencimento()
Dim rsVencimento As adodb.Recordset
Dim xString As String
    Set rsVencimento = New adodb.Recordset
    xString = "SELECT Codigo, [Dia 1 Inicial], [Dia 1 Final], [Dia 1 Vencimento], [Dia 2 Inicial], [Dia 2 Final], [Dia 2 Vencimento], [Dia 3 Inicial], [Dia 3 Final], [Dia 3 Vencimento] FROM Vencimento ORDER BY Codigo"
    Set rsVencimento = Conectar.RsConexao(xString)
    cbo_vencimento.Clear
    With rsVencimento
        If .RecordCount > 0 Then
            Do Until .EOF
                cbo_vencimento.AddItem rsVencimento("Codigo").Value & " - Dia " & Format(rsVencimento("Dia 1 Inicial").Value, "00") & " a " & Format(rsVencimento("Dia 1 Final").Value, "00") & " V. " & Format(rsVencimento("Dia 1 Vencimento").Value, "00") & "  |  Dia " & Format(rsVencimento("Dia 2 Inicial").Value, "00") & " a " & Format(rsVencimento("Dia 2 Final").Value, "00") & " V. " & Format(rsVencimento("Dia 2 Vencimento").Value, "00") & "  |  Dia " & Format(rsVencimento("Dia 3 Inicial").Value, "00") & " a " & Format(rsVencimento("Dia 3 Final").Value, "00") & " V. " & Format(rsVencimento("Dia 3 Vencimento").Value, "00")
                cbo_vencimento.ItemData(cbo_vencimento.NewIndex) = rsVencimento("Codigo").Value
                .MoveNext
            Loop
        End If
    End With
    rsVencimento.Close
    Set rsVencimento = Nothing
End Sub
Private Sub PreencheCboConvenio()
Dim rsConvenio As adodb.Recordset
Dim xString As String
    Set rsConvenio = New adodb.Recordset
    xString = "SELECT Codigo, Nome FROM Convenio ORDER BY Nome"
    Set rsConvenio = Conectar.RsConexao(xString)
    cbo_convenio.Clear
    With rsConvenio
        If .RecordCount > 0 Then
            Do Until .EOF
                cbo_convenio.AddItem rsConvenio("Nome").Value
                cbo_convenio.ItemData(cbo_convenio.NewIndex) = rsConvenio("Codigo").Value
                .MoveNext
            Loop
        End If
    End With
    rsConvenio.Close
    Set rsConvenio = Nothing
End Sub
Private Sub PreencheCboEmpresa()
Dim rsEmpresa As adodb.Recordset
Dim xString As String
    Set rsEmpresa = New adodb.Recordset
    xString = "SELECT Codigo, Nome FROM Empresas ORDER BY Nome"
    Set rsEmpresa = Conectar.RsConexao(xString)
    cbo_empresa.Clear
    With rsEmpresa
        If .RecordCount > 0 Then
            Do Until .EOF
                cbo_empresa.AddItem rsEmpresa("Nome").Value
                cbo_empresa.ItemData(cbo_empresa.NewIndex) = rsEmpresa("Codigo").Value
                .MoveNext
            Loop
        End If
    End With
    rsEmpresa.Close
    Set rsEmpresa = Nothing
End Sub
Private Sub msk_cep_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_telefone.SetFocus
    End If
End Sub
Private Sub msk_cgc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        chk_imprime_duplicata.SetFocus
    End If
End Sub
Private Sub msk_cgc_LostFocus()
    If Len(msk_cgc.Text) = 14 Then
        If Not CalculaDigitoCNPJ(msk_cgc.Text) Then
            msk_cgc.SetFocus
        End If
    End If
End Sub
Private Sub msk_cpf_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_identidade.SetFocus
    End If
End Sub
Private Sub msk_cpf_LostFocus()
    If Len(msk_cpf) = 11 Then
        If Not CalculaDigitoCPF(msk_cpf) Then
            msk_cpf.SetFocus
        End If
    End If
End Sub
Private Sub txtInscricaoEstadual_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        chk_inativo.SetFocus
    End If
End Sub
Private Sub txt_telefone_GotFocus()
    txt_telefone.Text = fDesmascaraTelefone(txt_telefone.Text)
End Sub
Private Sub txt_telefone_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_cpf.SetFocus
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
        txt_razao_social.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_codigo_LostFocus()
    If lOpcao = 1 And txt_codigo.Text <> "" Then
        If Cliente.LocalizarCodigo(CLng(txt_codigo.Text)) Then
            MsgBox "Já existe cliente cadastrado com este código." & Chr(10) & Chr(10) & "Mude o código informado.", vbInformation, "Duplicidade de Registro!"
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
Private Sub txt_identidade_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_orgao_emissor.SetFocus
    End If
End Sub
Private Sub txt_orgao_emissor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_cgc.SetFocus
    End If
End Sub
Private Sub txt_razao_social_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_endereco.SetFocus
    End If
End Sub
Private Sub txt_razao_social_LostFocus()
    If lOpcao = 1 And txt_razao_social.Text <> "" Then
        If Cliente.LocalizarNome(txt_razao_social.Text) Then
            If (MsgBox("Já existe cliente cadastrado com este nome." & Chr(10) & Chr(10) & "Deseja cadastrar assim mesmo?", 4 + 32 + 256, "Duplicidade de Registro!")) = 7 Then
                txt_razao_social.SetFocus
            End If
        End If
    End If
End Sub
Private Sub txt_telefone_LostFocus()
    txt_telefone.Text = fMascaraTelefone(txt_telefone.Text)
End Sub
Private Sub txt_uf_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_cep.SetFocus
    End If
End Sub
Private Sub txtEmail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtFormaPagamento.SetFocus
    End If
End Sub
Private Sub txtFormaPagamento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
