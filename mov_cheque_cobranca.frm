VERSION 5.00
Begin VB.Form movimento_cheque_cobranca 
   Caption         =   "Movimentação de Cheques em Cobrança"
   ClientHeight    =   4095
   ClientLeft      =   4185
   ClientTop       =   1875
   ClientWidth     =   6975
   Icon            =   "mov_cheque_cobranca.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "mov_cheque_cobranca.frx":030A
   ScaleHeight     =   4095
   ScaleWidth      =   6975
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   120
      Picture         =   "mov_cheque_cobranca.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Cria um novo registro."
      Top             =   3120
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   1020
      Picture         =   "mov_cheque_cobranca.frx":19E2
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Altera o registro atual."
      Top             =   3120
      Width           =   795
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   1920
      Picture         =   "mov_cheque_cobranca.frx":2EDC
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Exclui o registro atual."
      Top             =   3120
      Width           =   795
   End
   Begin VB.CommandButton cmd_pesquisa 
      Caption         =   "&Pesquisa"
      Height          =   855
      Left            =   2820
      Picture         =   "mov_cheque_cobranca.frx":456E
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Pesquisa um registro específico."
      Top             =   3120
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3720
      Picture         =   "mov_cheque_cobranca.frx":59E0
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   3120
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Enabled         =   0   'False
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      Begin VB.TextBox txt_conta 
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   6
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txt_banco 
         Height          =   285
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   2
         Top             =   240
         Width           =   435
      End
      Begin VB.TextBox txt_emitente 
         Height          =   300
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   12
         Top             =   2100
         Width           =   4935
      End
      Begin VB.TextBox txt_cheque 
         Height          =   300
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   8
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox txt_motivo 
         Height          =   300
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   14
         Top             =   2460
         Width           =   2835
      End
      Begin VB.TextBox msk_valor 
         Height          =   300
         Left            =   1680
         TabIndex        =   10
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txt_agencia 
         Height          =   285
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   4
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "&Número da Conta"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Número do &Banco"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Nome do &Emitente"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2100
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Número &do Cheque"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Número da &Agência"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "&Motivo"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2460
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "&Valor do Cheque"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   1455
      End
   End
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   4680
      TabIndex        =   22
      Top             =   3000
      Width           =   2175
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "mov_cheque_cobranca.frx":7072
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "mov_cheque_cobranca.frx":856C
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "mov_cheque_cobranca.frx":9A66
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "mov_cheque_cobranca.frx":AED8
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   5160
      Picture         =   "mov_cheque_cobranca.frx":C45A
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Confirma o registro atual."
      Top             =   3120
      Width           =   795
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   6060
      Picture         =   "mov_cheque_cobranca.frx":DA64
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Cancela o registro atual."
      Top             =   3120
      Width           =   795
   End
End
Attribute VB_Name = "movimento_cheque_cobranca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flag_movimento_cheque_cobranca As Integer
Dim lOpcao As Integer
Dim l_banco As Integer
Dim l_agencia As String
Dim l_conta As String
Dim l_cheque As String
Private ChequeCobranca As New cChequeCobranca

Private Sub AtualTabe()
    ChequeCobranca.Empresa = g_empresa
    ChequeCobranca.NumeroBanco = Val(txt_banco.Text)
    ChequeCobranca.NumeroAgencia = Val(txt_agencia.Text)
    ChequeCobranca.NumeroConta = txt_conta.Text
    ChequeCobranca.NumeroCheque = txt_cheque.Text
    ChequeCobranca.valor = fValidaValor2(msk_valor.Text)
    ChequeCobranca.Emitente = txt_emitente.Text
    ChequeCobranca.Motivo = txt_motivo.Text
End Sub
Private Sub AtualTela()
    l_banco = ChequeCobranca.NumeroBanco
    l_agencia = ChequeCobranca.NumeroAgencia
    l_conta = ChequeCobranca.NumeroConta
    l_cheque = ChequeCobranca.NumeroCheque
    txt_banco.Text = ChequeCobranca.NumeroBanco
    txt_agencia.Text = ChequeCobranca.NumeroAgencia
    txt_conta.Text = ChequeCobranca.NumeroConta
    txt_cheque.Text = ChequeCobranca.NumeroCheque
    msk_valor.Text = Format(ChequeCobranca.valor, "###,##0.00")
    txt_emitente.Text = ChequeCobranca.Emitente
    txt_motivo.Text = ChequeCobranca.Motivo
    frm_dados.Enabled = False
End Sub
Private Sub Finaliza()
    Set ChequeCobranca = Nothing
End Sub
Private Sub Inclui()
    lOpcao = 1
    DesativaBotoes
    cmd_novo.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
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
    If ChequeCobranca.LocalizarAnterior() Then
        AtualTela
    Else
        MsgBox "Início de Arquivo.", vbExclamation, "Atenção!"
        cmd_proximo.SetFocus
    End If
End Sub
Private Sub cmd_cancelar_Click()
    LimpaTela
    If ChequeCobranca.LocalizarUltimo(g_empresa) Then
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
    txt_banco = ""
    txt_agencia = ""
    txt_conta = ""
    txt_cheque = ""
    msk_valor = ""
    txt_emitente = ""
    txt_motivo = ""
End Sub
Private Sub cmd_excluir_Click()
    If Val(txt_banco.Text) > 0 Then
        If (MsgBox("Deseja Realmente Excluir Este Registro?", 4 + 32 + 256, "Exclusão de Registro!")) = 6 Then
            If ChequeCobranca.Excluir(g_empresa, l_cheque, l_banco, l_agencia, l_conta) Then
                LimpaTela
                If ChequeCobranca.LocalizarUltimo(g_empresa) Then
                    AtualTela
                    DesativaBotoes
                    cmd_novo.Enabled = True
                    cmd_sair.Enabled = True
                    cmd_novo.SetFocus
                End If
            Else
                MsgBox "Não foi possível excluir cheque em cobrança.", vbInformation, "Erro de Integridade!"
            End If
        End If
    End If
End Sub
Private Sub cmd_novo_Click()
    LimpaTela
    Inclui
    frm_dados.Enabled = True
    txt_banco.SetFocus
End Sub
Private Sub cmd_ok_Click()
    On Error GoTo FileError
    If ValidaCampos Then
        AtivaBotoes
        If lOpcao = 1 Then
            AtualTabe
            If ChequeCobranca.Incluir Then
                l_banco = ChequeCobranca.NumeroBanco
                l_agencia = ChequeCobranca.NumeroAgencia
                l_conta = ChequeCobranca.NumeroConta
                l_cheque = ChequeCobranca.NumeroCheque
            Else
                MsgBox "Não foi possível incluir cheque em cobrança.", vbInformation, "Erro de Integridade!"
            End If
        ElseIf lOpcao = 2 Then
            AtualTabe
            If ChequeCobranca.Alterar(g_empresa, l_cheque, l_banco, l_agencia, l_conta) Then
                l_banco = ChequeCobranca.NumeroBanco
                l_agencia = ChequeCobranca.NumeroAgencia
                l_conta = ChequeCobranca.NumeroConta
                l_cheque = ChequeCobranca.NumeroCheque
            Else
                MsgBox "Não foi possível alterar cheque em cobrança.", vbInformation, "Erro de Integridade!"
            End If
        End If
        If ChequeCobranca.LocalizarCodigo(g_empresa, l_cheque, l_banco, l_agencia, l_conta) Then
            AtualTela
        Else
            MsgBox "Não foi possível localizar cheque em cobrança.", vbInformation, "Erro de Integridade!"
        End If
        cmd_novo.SetFocus
    End If
    Exit Sub
FileError:
    Exit Sub
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If Not Val(txt_banco.Text) > 0 Then
        MsgBox "Informe o número do banco.", vbInformation, "Atenção!"
        txt_banco.SetFocus
    ElseIf Not Val(txt_agencia.Text) > 0 Then
        MsgBox "Informe o número da agencia.", vbInformation, "Atenção!"
        txt_agencia.SetFocus
    ElseIf Not Val(txt_conta.Text) > 0 Then
        MsgBox "Informe o número da conta.", vbInformation, "Atenção!"
        txt_conta.SetFocus
    ElseIf Not Val(txt_cheque.Text) > 0 Then
        MsgBox "Informe o número do cheque.", vbInformation, "Atenção!"
        txt_cheque.SetFocus
    ElseIf Not fValidaValor2(msk_valor.Text) > 0 Then
        MsgBox "Informe o valor do cheque.", vbInformation, "Atenção!"
        msk_valor.SetFocus
    ElseIf Not txt_emitente.Text <> "" Then
        MsgBox "Informe o nome do emitente.", vbInformation, "Atenção!"
        txt_emitente.SetFocus
    ElseIf Not txt_motivo.Text <> "" Then
        MsgBox "Informe o motivo.", vbInformation, "Atenção!"
        txt_motivo.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_pesquisa_Click()
    consulta_cheque_cobranca.Show 1
    If Len(g_string) > 0 Then
        l_cheque = RetiraGString(1)
        l_banco = RetiraGString(2)
        l_agencia = RetiraGString(3)
        l_conta = RetiraGString(4)
        If ChequeCobranca.LocalizarCodigo(g_empresa, l_cheque, l_banco, l_agencia, l_conta) Then
            AtualTela
        Else
            MsgBox "Não foi possível localizar cheque em cobrança.", vbInformation, "Erro de Integridade!"
        End If
    End If
End Sub
Private Sub cmd_primeiro_Click()
    If ChequeCobranca.LocalizarPrimeiro Then
        AtualTela
        cmd_proximo.SetFocus
    Else
        MsgBox "Não há registros nesta empresa.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    If ChequeCobranca.LocalizarProximo Then
        AtualTela
    Else
        MsgBox "Fim de Arquivo.", vbExclamation, "Atenção!"
        cmd_anterior.SetFocus
    End If
End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub cmd_ultimo_Click()
    If ChequeCobranca.LocalizarUltimo(g_empresa) Then
        AtualTela
        cmd_anterior.SetFocus
    Else
        MsgBox "Não há registros nesta empresa.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub Form_Activate()
    If flag_movimento_cheque_cobranca = 0 Then
        DesativaBotoes
        If ChequeCobranca.LocalizarUltimo(g_empresa) Then
            AtualTela
            AtivaBotoes
        Else
            cmd_novo.Enabled = True
            cmd_sair.Enabled = True
        End If
        cmd_novo.SetFocus
    Else
        flag_movimento_cheque_cobranca = 0
    End If
End Sub
Private Sub AtivaBotoes()
    cmd_novo.Enabled = True
    cmd_alterar.Enabled = True
    cmd_excluir.Enabled = True
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
Private Sub Form_Deactivate()
    flag_movimento_cheque_cobranca = 1
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
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_valor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_emitente.SetFocus
    End If
End Sub
Private Sub msk_valor_LostFocus()
    If Val(msk_valor.Text) > 0 Then
        msk_valor.Text = Format(msk_valor.Text, "###,##0.00")
    End If
End Sub
Private Sub txt_agencia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_conta.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_banco_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_agencia.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_cheque_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_valor.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_cheque_LostFocus()
    If lOpcao = 1 Then
        If ChequeCobranca.LocalizarCodigo(g_empresa, txt_cheque.Text, Val(txt_banco.Text), txt_agencia.Text, txt_conta.Text) Then
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
Private Sub txt_emitente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_motivo.SetFocus
    End If
End Sub
Private Sub txt_motivo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
