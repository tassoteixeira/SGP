VERSION 5.00
Begin VB.Form cadastro_conversao_medicao 
   Caption         =   "Tabela de Conversão de Medição de Tanques"
   ClientHeight    =   3195
   ClientLeft      =   2325
   ClientTop       =   3105
   ClientWidth     =   7050
   Icon            =   "cadastro_conversao_medicao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "cadastro_conversao_medicao.frx":030A
   ScaleHeight     =   3195
   ScaleWidth      =   7050
   Begin VB.CommandButton cmd_pesquisa 
      Caption         =   "&Pesquisa"
      Height          =   855
      Left            =   2820
      Picture         =   "cadastro_conversao_medicao.frx":0750
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Pesquisa um registro específico."
      Top             =   2220
      Width           =   795
   End
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   120
      Picture         =   "cadastro_conversao_medicao.frx":1BC2
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Cria um novo registro."
      Top             =   2220
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   1020
      Picture         =   "cadastro_conversao_medicao.frx":3254
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Altera o registro atual."
      Top             =   2220
      Width           =   795
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   1920
      Picture         =   "cadastro_conversao_medicao.frx":474E
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Exclui o registro atual."
      Top             =   2220
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3720
      Picture         =   "cadastro_conversao_medicao.frx":5DE0
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   2220
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Enabled         =   0   'False
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   6795
      Begin VB.TextBox txt_medicao_tanque_30 
         Height          =   285
         Left            =   2880
         MaxLength       =   6
         TabIndex        =   10
         Top             =   1620
         Width           =   1095
      End
      Begin VB.TextBox txt_medida 
         Height          =   285
         Left            =   2880
         MaxLength       =   3
         TabIndex        =   2
         Top             =   180
         Width           =   495
      End
      Begin VB.TextBox txt_medicao_tanque_20 
         Height          =   285
         Left            =   2880
         MaxLength       =   6
         TabIndex        =   8
         Top             =   1260
         Width           =   1095
      End
      Begin VB.TextBox txt_medicao_tanque_15 
         Height          =   285
         Left            =   2880
         MaxLength       =   6
         TabIndex        =   6
         Top             =   900
         Width           =   1095
      End
      Begin VB.TextBox txt_medicao_tanque_10 
         Height          =   285
         Left            =   2880
         MaxLength       =   6
         TabIndex        =   4
         Top             =   540
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "&Medição do tanque de 30.000 lts."
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   1620
         Width           =   2715
      End
      Begin VB.Label Label2 
         Caption         =   "&Medida"
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   180
         Width           =   2715
      End
      Begin VB.Label Label9 
         Caption         =   "&Medição do tanque de 20.000 lts."
         Height          =   285
         Index           =   14
         Left            =   120
         TabIndex        =   7
         Top             =   1260
         Width           =   2715
      End
      Begin VB.Label Label6 
         Caption         =   "&Medição do tanque de 15.000 lts."
         Height          =   285
         Index           =   10
         Left            =   120
         TabIndex        =   5
         Top             =   900
         Width           =   2715
      End
      Begin VB.Label Label2 
         Caption         =   "&Medição do tanque de 10.000 lts."
         Height          =   285
         Index           =   6
         Left            =   120
         TabIndex        =   3
         Top             =   540
         Width           =   2715
      End
   End
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   4740
      TabIndex        =   20
      Top             =   2100
      Width           =   2175
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "cadastro_conversao_medicao.frx":7472
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "cadastro_conversao_medicao.frx":896C
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "cadastro_conversao_medicao.frx":9E66
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "cadastro_conversao_medicao.frx":B2D8
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   5220
      Picture         =   "cadastro_conversao_medicao.frx":C85A
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Confirma o registro atual."
      Top             =   2220
      Width           =   795
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   6120
      Picture         =   "cadastro_conversao_medicao.frx":DE64
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Cancela o registro atual."
      Top             =   2220
      Width           =   795
   End
End
Attribute VB_Name = "cadastro_conversao_medicao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lFlagCadastro As Integer
Dim lOpcao As Integer
Dim lMedida As Integer
Private ConversaoMedicaoCombustivel As New cConvMedicaoComb
Private Sub AtualTabe()
    ConversaoMedicaoCombustivel.Empresa = g_empresa
    ConversaoMedicaoCombustivel.Medida = Val(txt_medida.Text)
    ConversaoMedicaoCombustivel.MedicaoTanque10 = CLng(txt_medicao_tanque_10.Text)
    ConversaoMedicaoCombustivel.MedicaoTanque15 = CLng(txt_medicao_tanque_15.Text)
    ConversaoMedicaoCombustivel.MedicaoTanque20 = CLng(txt_medicao_tanque_20.Text)
    ConversaoMedicaoCombustivel.MedicaoTanque30 = CLng(txt_medicao_tanque_30.Text)
End Sub
Private Sub AtualTela()
    lMedida = ConversaoMedicaoCombustivel.Medida
    txt_medida.Text = Format(ConversaoMedicaoCombustivel.Medida, "##0")
    txt_medicao_tanque_10.Text = Format(ConversaoMedicaoCombustivel.MedicaoTanque10, "##,##0")
    txt_medicao_tanque_15.Text = Format(ConversaoMedicaoCombustivel.MedicaoTanque15, "##,##0")
    txt_medicao_tanque_20.Text = Format(ConversaoMedicaoCombustivel.MedicaoTanque20, "##,##0")
    txt_medicao_tanque_30.Text = Format(ConversaoMedicaoCombustivel.MedicaoTanque30, "##,##0")
    frm_dados.Enabled = False
End Sub
Private Sub Finaliza()
    Set ConversaoMedicaoCombustivel = Nothing
    frm_cadastro.Show
End Sub
Private Sub Inclui()
    lOpcao = 1
    DesativaBotoes
    cmd_novo.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    txt_medida.Text = 1
    If ConversaoMedicaoCombustivel.LocalizarUltimo(g_empresa) Then
        txt_medida.Text = ConversaoMedicaoCombustivel.Medida + 1
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
    txt_medicao_tanque_10.SetFocus
End Sub
Private Sub cmd_anterior_Click()
    If ConversaoMedicaoCombustivel.LocalizarAnterior Then
        AtualTela
    Else
        MsgBox "Início de Arquivo.", vbInformation, "Atenção!"
        cmd_proximo.SetFocus
    End If
End Sub
Private Sub cmd_cancelar_Click()
    LimpaTela
    If ConversaoMedicaoCombustivel.LocalizarCodigo(g_empresa, lMedida) Then
        AtualTela
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
    txt_medida.Text = ""
    txt_medicao_tanque_10.Text = ""
    txt_medicao_tanque_15.Text = ""
    txt_medicao_tanque_20.Text = ""
    txt_medicao_tanque_30.Text = ""
End Sub
Private Sub cmd_excluir_Click()
    If Val(txt_medida.Text) > 0 Then
        If (MsgBox("Deseja Realmente Excluir Este Registro?", 4 + 32 + 256, "Exclusão de Registro!")) = 6 Then
            If ConversaoMedicaoCombustivel.Excluir(g_empresa, Val(txt_medida.Text)) Then
                LimpaTela
                If ConversaoMedicaoCombustivel.LocalizarUltimo(g_empresa) Then
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
    LimpaTela
    Inclui
    frm_dados.Enabled = True
    txt_medicao_tanque_10.SetFocus
End Sub
Private Sub cmd_ok_Click()
    On Error GoTo FileError
    If ValidaCampos Then
        AtivaBotoes
        If lOpcao = 1 Then
            AtualTabe
            If ConversaoMedicaoCombustivel.Incluir Then
                lMedida = Val(txt_medida.Text)
            Else
                MsgBox "Não foi possível incluir este registro!", vbInformation, "Erro de Verificação!"
            End If
        ElseIf lOpcao = 2 Then
            AtualTabe
            If Not ConversaoMedicaoCombustivel.Alterar(g_empresa, lMedida) Then
                MsgBox "Não foi possível alterar este registro!", vbInformation, "Erro de Verificação!"
            End If
        End If
        lOpcao = 0
        Call ConversaoMedicaoCombustivel.LocalizarCodigo(g_empresa, lMedida)
        AtualTela
        cmd_novo.SetFocus
    End If
    Exit Sub
FileError:
    MsgBox Error
    'ErroArquivo tbl_movimento_cheque.Name, "Chequeo"
    Exit Sub
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If Not Val(txt_medida.Text) > 0 Then
        MsgBox "Informe o número da medida.", vbInformation, "Atenção!"
        txt_medida.SetFocus
    ElseIf txt_medicao_tanque_10.Text = "" Then
        MsgBox "Informe a medição do tanque de 10.000 litros.", vbInformation, "Atenção!"
        txt_medicao_tanque_10.SetFocus
    ElseIf txt_medicao_tanque_15.Text = "" Then
        MsgBox "Informe a medição do tanque de 15.000 litros.", vbInformation, "Atenção!"
        txt_medicao_tanque_15.SetFocus
    ElseIf txt_medicao_tanque_20.Text = "" Then
        MsgBox "Informe a medição do tanque de 20.000 litros.", vbInformation, "Atenção!"
        txt_medicao_tanque_20.SetFocus
    ElseIf txt_medicao_tanque_30.Text = "" Then
        MsgBox "Informe a medição do tanque de 30.000 litros.", vbInformation, "Atenção!"
        txt_medicao_tanque_30.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_pesquisa_Click()
    consulta_conversao_medicao.Show 1
    If Len(g_string) > 0 Then
        lMedida = RetiraGString(1)
        Call ConversaoMedicaoCombustivel.LocalizarCodigo(g_empresa, lMedida)
        AtualTela
    End If
End Sub
Private Sub cmd_primeiro_Click()
    If ConversaoMedicaoCombustivel.LocalizarPrimeiro Then
        AtualTela
        cmd_proximo.SetFocus
    Else
        MsgBox "Não há registros.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    If ConversaoMedicaoCombustivel.LocalizarProximo Then
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
    If ConversaoMedicaoCombustivel.LocalizarUltimo(g_empresa) Then
        AtualTela
        cmd_anterior.SetFocus
    Else
        MsgBox "Não há registros.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub Form_Activate()
    If lFlagCadastro = 0 Then
        lOpcao = 0
        DesativaBotoes
        If ConversaoMedicaoCombustivel.LocalizarUltimo(g_empresa) Then
            AtualTela
            AtivaBotoes
        Else
            cmd_novo.Enabled = True
            cmd_sair.Enabled = True
        End If
        If cmd_novo.Enabled Then
            cmd_novo.SetFocus
        End If
    Else
        lFlagCadastro = 0
    End If
    Screen.MousePointer = 1
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
    CentraForm Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
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
Private Sub txt_medicao_tanque_10_GotFocus()
    txt_medicao_tanque_10.SelStart = 0
    txt_medicao_tanque_10.SelLength = Len(txt_medicao_tanque_10.Text)
End Sub
Private Sub txt_medicao_tanque_10_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_medicao_tanque_15.SetFocus
    End If
End Sub
Private Sub txt_medicao_tanque_10_LostFocus()
    If txt_medicao_tanque_10.Text = "" Then
        txt_medicao_tanque_10.Text = 0
    End If
    txt_medicao_tanque_10.Text = Format(txt_medicao_tanque_10.Text, "##,##0")
End Sub
Private Sub txt_medicao_tanque_15_GotFocus()
    txt_medicao_tanque_15.SelStart = 0
    txt_medicao_tanque_15.SelLength = Len(txt_medicao_tanque_15.Text)
End Sub
Private Sub txt_medicao_tanque_15_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_medicao_tanque_20.SetFocus
    End If
End Sub
Private Sub txt_medicao_tanque_15_LostFocus()
    If txt_medicao_tanque_15.Text = "" Then
        txt_medicao_tanque_15.Text = 0
    End If
    txt_medicao_tanque_15.Text = Format(txt_medicao_tanque_15.Text, "##,##0")
End Sub
Private Sub txt_medicao_tanque_20_GotFocus()
    txt_medicao_tanque_20.SelStart = 0
    txt_medicao_tanque_20.SelLength = Len(txt_medicao_tanque_20.Text)
End Sub
Private Sub txt_medicao_tanque_20_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_medicao_tanque_30.SetFocus
    End If
End Sub
Private Sub txt_medicao_tanque_20_LostFocus()
    If txt_medicao_tanque_20.Text = "" Then
        txt_medicao_tanque_20.Text = 0
    End If
    txt_medicao_tanque_20.Text = Format(txt_medicao_tanque_20.Text, "##,##0")
End Sub
Private Sub txt_medicao_tanque_30_GotFocus()
    txt_medicao_tanque_30.SelStart = 0
    txt_medicao_tanque_30.SelLength = Len(txt_medicao_tanque_30.Text)
End Sub
Private Sub txt_medicao_tanque_30_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub txt_medicao_tanque_30_LostFocus()
    If txt_medicao_tanque_30.Text = "" Then
        txt_medicao_tanque_30.Text = 0
    End If
    txt_medicao_tanque_30.Text = Format(txt_medicao_tanque_30.Text, "##,##0")
End Sub
Private Sub txt_medida_GotFocus()
    txt_medida.SelStart = 0
    txt_medida.SelLength = Len(txt_medida.Text)
End Sub
Private Sub txt_medida_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_medicao_tanque_10.SetFocus
    End If
End Sub
Private Sub txt_medida_LostFocus()
    txt_medida.Text = Format(txt_medida.Text, "##0")
End Sub
