VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form movimento_advertencia_suspencao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimento de Advertência / Suspenção"
   ClientHeight    =   4440
   ClientLeft      =   1305
   ClientTop       =   570
   ClientWidth     =   6930
   Icon            =   "mov_advertencia_suspencao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "mov_advertencia_suspencao.frx":030A
   ScaleHeight     =   4440
   ScaleWidth      =   6930
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   120
      Picture         =   "mov_advertencia_suspencao.frx":0750
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Cria um novo registro."
      Top             =   3480
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   1020
      Picture         =   "mov_advertencia_suspencao.frx":1DE2
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Altera o registro atual."
      Top             =   3480
      Width           =   795
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   1920
      Picture         =   "mov_advertencia_suspencao.frx":32DC
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Exclui o registro atual."
      Top             =   3480
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Caption         =   "&Sair"
      Height          =   855
      Left            =   2820
      Picture         =   "mov_advertencia_suspencao.frx":496E
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   3480
      Width           =   795
   End
   Begin VB.Frame frmDados 
      Enabled         =   0   'False
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6675
      Begin VB.TextBox txt_motivo 
         Height          =   315
         Index           =   3
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   14
         Top             =   2820
         Width           =   4815
      End
      Begin VB.TextBox txt_motivo 
         Height          =   315
         Index           =   2
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   13
         Top             =   2460
         Width           =   4815
      End
      Begin VB.TextBox txt_motivo 
         Height          =   315
         Index           =   1
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   12
         Top             =   2100
         Width           =   4815
      End
      Begin VB.TextBox txt_dia 
         Height          =   315
         Left            =   6060
         MaxLength       =   2
         TabIndex        =   9
         Top             =   1260
         Width           =   435
      End
      Begin VB.Data dta_funcionario 
         Caption         =   "dta_funcionario"
         Connect         =   "Access"
         DatabaseName    =   "Sgp_data.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   4080
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   360
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox txt_motivo 
         Height          =   315
         Index           =   0
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   11
         Top             =   1740
         Width           =   4815
      End
      Begin VB.Frame frmTipo 
         Caption         =   "Advertência ou Suspensão"
         Height          =   555
         Left            =   180
         TabIndex        =   5
         Top             =   1080
         Width           =   3555
         Begin VB.OptionButton optTipo 
            Caption         =   "&Suspensão"
            Height          =   255
            Index           =   1
            Left            =   1860
            TabIndex        =   7
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton optTipo 
            Caption         =   "&Advertência"
            Height          =   255
            Index           =   0
            Left            =   300
            TabIndex        =   6
            Top             =   240
            Width           =   1455
         End
      End
      Begin MSDBCtls.DBCombo dbcbo_funcionario 
         Bindings        =   "mov_advertencia_suspencao.frx":6000
         Height          =   315
         Left            =   1680
         TabIndex        =   4
         Top             =   660
         Width           =   4815
         _ExtentX        =   8493
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
      Begin MSMask.MaskEdBox msk_data 
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
      Begin VB.Label Label1 
         Caption         =   "Dias"
         Height          =   255
         Index           =   3
         Left            =   5640
         TabIndex        =   8
         Top             =   1320
         Width           =   435
      End
      Begin VB.Label Label1 
         Caption         =   "Motivo"
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   10
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Funcionário"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   3
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Data"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   1
         Top             =   300
         Width           =   1455
      End
   End
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   4680
      TabIndex        =   21
      Top             =   3360
      Width           =   2175
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "mov_advertencia_suspencao.frx":601E
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "mov_advertencia_suspencao.frx":75A0
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "mov_advertencia_suspencao.frx":8A12
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "mov_advertencia_suspencao.frx":9F0C
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   5160
      Picture         =   "mov_advertencia_suspencao.frx":B406
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Confirma o registro atual."
      Top             =   3480
      Width           =   795
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   6060
      Picture         =   "mov_advertencia_suspencao.frx":CA10
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Cancela o registro atual."
      Top             =   3480
      Width           =   795
   End
End
Attribute VB_Name = "movimento_advertencia_suspencao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lOpcao As Integer
Dim lData As Date
Dim lFuncionario As Integer
Dim l_movimento_advertencia_suspencao As Integer
Private MovAdvertenciaSuspencao As New cMovAdvertenciaSuspencao
Private Sub AtualTabe()
    MovAdvertenciaSuspencao.Empresa = g_empresa
    MovAdvertenciaSuspencao.Data = Format(msk_data.Text, "dd/mm/yyyy")
    MovAdvertenciaSuspencao.CodigoFuncionario = Val(dbcbo_funcionario.BoundText)
    If optTipo(0) Then
        MovAdvertenciaSuspencao.AdvertenciaouSuspencao = "A"
    Else
        MovAdvertenciaSuspencao.AdvertenciaouSuspencao = "S"
    End If
    MovAdvertenciaSuspencao.dia = Val(txt_dia.Text)
    MovAdvertenciaSuspencao.Motivo1 = "" & txt_motivo(0).Text
    MovAdvertenciaSuspencao.Motivo2 = "" & txt_motivo(1).Text
    MovAdvertenciaSuspencao.Motivo3 = "" & txt_motivo(2).Text
    MovAdvertenciaSuspencao.Motivo4 = "" & txt_motivo(3).Text
End Sub
Private Sub AtualTela()
    lFuncionario = MovAdvertenciaSuspencao.CodigoFuncionario
    lData = MovAdvertenciaSuspencao.Data
    msk_data.Text = Format(MovAdvertenciaSuspencao.Data, "dd/mm/yyyy")
    dbcbo_funcionario.BoundText = MovAdvertenciaSuspencao.CodigoFuncionario
    If MovAdvertenciaSuspencao.AdvertenciaouSuspencao = "A" Then
        optTipo(0).Value = True
    Else
        optTipo(1).Value = True
    End If
    txt_dia.Text = Format(MovAdvertenciaSuspencao.dia, "##")
    txt_motivo(0).Text = MovAdvertenciaSuspencao.Motivo1
    txt_motivo(1).Text = MovAdvertenciaSuspencao.Motivo2
    txt_motivo(2).Text = MovAdvertenciaSuspencao.Motivo3
    txt_motivo(3).Text = MovAdvertenciaSuspencao.Motivo4
    frmDados.Enabled = False
End Sub
Private Sub Finaliza()
    Set MovAdvertenciaSuspencao = Nothing
End Sub
Private Sub Incluir()
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
    frmDados.Enabled = True
    If optTipo(0) Then
        optTipo(0).SetFocus
    Else
        optTipo(1).SetFocus
    End If
End Sub
Private Sub cmd_anterior_Click()
    If MovAdvertenciaSuspencao.LocalizarAnterior Then
        AtualTela
    Else
        MsgBox "Início de Arquivo.", 48, "Atenção!"
        cmd_proximo.SetFocus
    End If
End Sub
Private Sub cmd_cancelar_Click()
    LimpaTela
    If MovAdvertenciaSuspencao.LocalizarCodigo(g_empresa, lData, lFuncionario) Then
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
    msk_data.Text = "__/__/____"
    dbcbo_funcionario.BoundText = ""
    optTipo(0).Value = True
    txt_dia.Text = ""
    txt_motivo(0).Text = ""
    txt_motivo(1).Text = ""
    txt_motivo(2).Text = ""
    txt_motivo(3).Text = ""
End Sub
Private Sub cmd_cancelar_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    cmd_cancelar.Value = True
End Sub
Private Sub cmd_excluir_Click()
    If IsDate(msk_data) Then
        If (MsgBox("Deseja Realmente Excluir Este Registro?", 4 + 32 + 256, "Exclusão de Registro!")) = 6 Then
            If Not MovAdvertenciaSuspencao.Excluir(g_empresa, CDate(lData), Val(Me.dbcbo_funcionario.BoundText)) Then
                MsgBox "Não foi possível excluir este registro.", vbInformation, "Erro de Integridade"
            End If
            LimpaTela
            If MovAdvertenciaSuspencao.LocalizarUltimo(g_empresa) Then
                AtualTela
            Else
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
    Incluir
    frmDados.Enabled = True
    msk_data.SetFocus
End Sub
Private Sub cmd_ok_Click()
    On Error GoTo FileError
    If ValidaCampos Then
        AtivaBotoes
        If lOpcao = 1 Then
            AtualTabe
            If MovAdvertenciaSuspencao.Incluir Then
                lFuncionario = Val(dbcbo_funcionario.BoundText)
                lData = CDate(msk_data.Text)
            Else
                MsgBox "Não foi possível incluir este registro!", vbInformation, "Erro de Integridade"
            End If
        ElseIf lOpcao = 2 Then
            AtualTabe
            If MovAdvertenciaSuspencao.Alterar(g_empresa, lData, lFuncionario) Then
                lFuncionario = Val(dbcbo_funcionario.BoundText)
                lData = CDate(msk_data.Text)
            Else
                MsgBox "Não foi possível alterar este registro!", vbInformation, "Erro de Integridade"
            End If
        End If
        If MovAdvertenciaSuspencao.LocalizarCodigo(g_empresa, lData, lFuncionario) Then
            AtualTela
            lOpcao = 0
            cmd_novo.SetFocus
        End If
    End If
    Exit Sub
FileError:
    MsgBox "Erro interno", vbInformation, "Erro"
    Exit Sub
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If Not IsDate(msk_data) Then
        MsgBox "Informe a data.", vbExclamation, "Atenção!"
        msk_data.SetFocus
    ElseIf Val(dbcbo_funcionario.BoundText) = 0 Then
        MsgBox "Escolha o funcionário.", 64, "Atenção!"
        dbcbo_funcionario.SetFocus
    ElseIf optTipo(0) = False And optTipo(1) = False Then
        MsgBox "Escolha o tipo de falta.", 64, "Atenção!"
        optTipo(0).SetFocus
    ElseIf optTipo(1) = True And Val(txt_dia) = 0 Then
        MsgBox "Informe os dias de suspenção.", 64, "Atenção!"
        txt_dia.SetFocus
    ElseIf optTipo(0) = True And Val(txt_dia) > 0 Then
        MsgBox "Advertência não pode ter dias.", 64, "Atenção!"
        txt_dia.SetFocus
    ElseIf txt_motivo(0) = "" Then
        MsgBox "Informe o motivo.", 64, "Atenção!"
        txt_motivo(0).SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_primeiro_Click()
    If MovAdvertenciaSuspencao.LocalizarPrimeiro Then
        AtualTela
        cmd_proximo.SetFocus
    Else
        MsgBox "Não há registros nesta empresa.", 64, "Erro de Verificação!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    If MovAdvertenciaSuspencao.LocalizarProximo Then
        AtualTela
    Else
        MsgBox "Fim de Arquivo.", 48, "Atenção!"
        cmd_anterior.SetFocus
    End If
End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub cmd_ultimo_Click()
    If MovAdvertenciaSuspencao.LocalizarUltimo(g_empresa) Then
        AtualTela
        cmd_anterior.SetFocus
    Else
        MsgBox "Não há registros nesta empresa.", 64, "Erro de Verificação!"
    End If
End Sub
Private Sub dbcbo_funcionario_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        optTipo(0).SetFocus
    End If
End Sub
Private Sub dbcbo_funcionario_LostFocus()
    If Val(dbcbo_funcionario.BoundText) > 0 Then
        If MovAdvertenciaSuspencao.LocalizarCodigo(g_empresa, CDate(msk_data.Text), Val(dbcbo_funcionario.BoundText)) Then
            MsgBox "Registro jâ Cadastrado", vbExclamation, "Atenção!"
            dbcbo_funcionario.SetFocus
        End If
    End If
End Sub
Private Sub Form_Activate()
    If l_movimento_advertencia_suspencao = 0 Then
        DesativaBotoes
        If MovAdvertenciaSuspencao.LocalizarUltimo(g_empresa) Then
            AtualTela
            AtivaBotoes
        Else
            cmd_novo.Enabled = True
            cmd_sair.Enabled = True
        End If
        cmd_novo.SetFocus
    Else
        l_movimento_advertencia_suspencao = 0
    End If
End Sub
Private Sub AtivaBotoes()
    cmd_novo.Enabled = True
    cmd_excluir.Enabled = True
    cmd_alterar.Enabled = True
    cmd_sair.Enabled = True
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
    frm_move.Visible = True
End Sub
Private Sub DesativaBotoes()
    cmd_novo.Enabled = False
    cmd_excluir.Enabled = False
    cmd_alterar.Enabled = False
    cmd_sair.Enabled = False
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
    frm_move.Visible = False
End Sub
Private Sub Form_Deactivate()
    l_movimento_advertencia_suspencao = 1
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
Private Sub msk_data_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dbcbo_funcionario.SetFocus
    End If
End Sub
Private Sub optTipo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_dia.SetFocus
    End If
End Sub
Private Sub txt_dia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_motivo(0).SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_motivo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Index < 2 Then
            txt_motivo(Index + 1).SetFocus
        Else
            cmd_ok.SetFocus
        End If
    End If
End Sub
