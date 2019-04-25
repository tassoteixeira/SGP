VERSION 5.00
Begin VB.Form cadastro_programa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Programas"
   ClientHeight    =   4575
   ClientLeft      =   2805
   ClientTop       =   3600
   ClientWidth     =   7035
   Icon            =   "cad_programa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4575
   ScaleWidth      =   7035
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   120
      Picture         =   "cad_programa.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Cria um novo registro."
      Top             =   3600
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   1020
      Picture         =   "cad_programa.frx":199C
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Altera o registro atual."
      Top             =   3600
      Width           =   795
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   1920
      Picture         =   "cad_programa.frx":2E96
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Exclui o registro atual."
      Top             =   3600
      Width           =   795
   End
   Begin VB.CommandButton cmd_pesquisa 
      Caption         =   "&Pesquisa"
      Height          =   855
      Left            =   2820
      Picture         =   "cad_programa.frx":4528
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Pesquisa um registro específico."
      Top             =   3600
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3720
      Picture         =   "cad_programa.frx":599A
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   3600
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Enabled         =   0   'False
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6795
      Begin VB.TextBox txt_configuravel 
         Height          =   285
         Left            =   1740
         MaxLength       =   1
         TabIndex        =   12
         Top             =   2940
         Width           =   255
      End
      Begin VB.TextBox txt_observacao 
         Height          =   1155
         Left            =   1740
         MaxLength       =   40
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   8
         Top             =   1320
         Width           =   4935
      End
      Begin VB.TextBox txt_disco 
         Height          =   285
         Left            =   1740
         MaxLength       =   40
         TabIndex        =   6
         Top             =   960
         Width           =   4935
      End
      Begin VB.TextBox txt_interno 
         Height          =   285
         Left            =   1740
         MaxLength       =   40
         TabIndex        =   4
         Top             =   600
         Width           =   4935
      End
      Begin VB.ComboBox cbo_tipo 
         Height          =   300
         Left            =   1740
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2580
         Width           =   615
      End
      Begin VB.TextBox txt_menu 
         Height          =   285
         Left            =   1740
         MaxLength       =   40
         TabIndex        =   2
         Top             =   300
         Width           =   4935
      End
      Begin VB.Label Label2 
         Caption         =   "Observação"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Nome Em Disco"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Nome Interno"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Tipo Programa"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2580
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Configurável"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   2940
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Nome P/ Menu"
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
      TabIndex        =   20
      Top             =   3480
      Width           =   2175
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "cad_programa.frx":702C
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "cad_programa.frx":85AE
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "cad_programa.frx":9A20
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "cad_programa.frx":AF1A
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   6120
      Picture         =   "cad_programa.frx":C414
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Cancela o registro atual."
      Top             =   3600
      Width           =   795
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   5220
      Picture         =   "cad_programa.frx":D90E
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Confirma o registro atual."
      Top             =   3600
      Width           =   795
   End
End
Attribute VB_Name = "cadastro_programa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lFlagCadastro As Integer
Dim lOpcao As Integer
Dim lNomeMenu As String
Dim lNomeDisco As String
Dim lCodigo As Integer
Private Programa As New cPrograma
Private Sub AtivaBotoes()
    cmd_novo.Enabled = True
    cmd_alterar.Enabled = True
    cmd_pesquisa.Enabled = True
    cmd_excluir.Enabled = True
    cmd_sair.Enabled = True
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
    frm_move.Visible = True
End Sub
Private Sub AtualTabe()
    If lOpcao = 1 Then
        Programa.Codigo = lCodigo
    End If
    Programa.NomeparaMenu = txt_menu.Text
    Programa.NomeInterno = txt_interno.Text
    Programa.NomenoDisco = txt_disco.Text
    Programa.Observacao = txt_observacao.Text
    Programa.Tipo = cbo_tipo.Text
    Programa.Configuravel = txt_configuravel.Text
End Sub
Private Sub AtualTela()
Dim i As Integer
    lNomeDisco = Programa.NomenoDisco
    lCodigo = Programa.Codigo
    txt_menu.Text = Programa.NomeparaMenu
    txt_interno.Text = Programa.NomeInterno
    txt_disco.Text = Programa.NomenoDisco
    txt_observacao.Text = Programa.Observacao
    For i = 0 To cbo_tipo.ListCount - 1
        cbo_tipo.ListIndex = i
        If cbo_tipo.Text = Programa.Tipo Then
            Exit For
        End If
    Next
    If cbo_tipo.Text <> Programa.Tipo Then
        cbo_tipo.ListIndex = -1
    End If
    txt_configuravel.Text = Programa.Configuravel
    frm_dados.Enabled = False
End Sub
Private Sub DesativaBotoes()
    cmd_novo.Enabled = False
    cmd_alterar.Enabled = False
    cmd_pesquisa.Enabled = False
    cmd_excluir.Enabled = False
    cmd_sair.Enabled = False
    frm_move.Visible = False
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
End Sub
Private Sub Finaliza()
    Set Programa = Nothing
    frm_cadastro.Show
End Sub
Private Sub Inclui()
    lOpcao = 1
    DesativaBotoes
    cmd_novo.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    lCodigo = Programa.ProximoCodigo
End Sub
Private Sub cbo_tipo_GotFocus()
    SendMessageLong cbo_tipo.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_tipo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_configuravel.SetFocus
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
    txt_menu.SetFocus
End Sub
Private Sub cmd_anterior_Click()
    If Programa.LocalizarAnterior Then
        AtualTela
    Else
        MsgBox "Início de Arquivo.", vbInformation, "Atenção!"
        cmd_proximo.SetFocus
    End If
End Sub
Private Sub cmd_cancelar_Click()
    LimpaTela
    lOpcao = 0
    If Programa.LocalizarNomeDisco(lNomeDisco) Then
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
    'txt_menu = ""
    'txt_interno = ""
    'txt_disco = ""
    'txt_observacao = ""
    'cbo_tipo.ListIndex = -1
    'txt_configuravel = ""
End Sub
Private Sub cmd_excluir_Click()
    If lNomeDisco <> "" Then
        If (MsgBox("Deseja Realmente Excluir Este Registro?", 4 + 32 + 256, "Exclusão de Registro!")) = 6 Then
            If Programa.Excluir(txt_disco.Text) Then
                LimpaTela
                If Programa.LocalizarUltimo Then
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
    txt_menu.SetFocus
End Sub
Private Sub cmd_ok_Click()
    On Error GoTo FileError
    If ValidaCampos Then
        AtivaBotoes
        If lOpcao = 1 Then
            AtualTabe
            If Programa.Incluir Then
                lNomeDisco = Val(txt_disco.Text)
            Else
                MsgBox "Não foi possível incluir este registro!", vbInformation, "Erro de Verificação!"
            End If
        ElseIf lOpcao = 2 Then
            AtualTabe
            If Not Programa.Alterar(lNomeDisco) Then
                MsgBox "Não foi possível alterar este registro!", vbInformation, "Erro de Verificação!"
            End If
        End If
        lOpcao = 0
        Call Programa.LocalizarNomeDisco(lNomeDisco)
        AtualTela
        cmd_novo.SetFocus
    End If
    Exit Sub
FileError:
    MsgBox Error
    'ErroArquivo tbl_programa.Name, "Programao"
    Exit Sub
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If txt_menu = "" Then
        MsgBox "Informe o nome p/ menu.", vbInformation, "Atenção!"
        txt_menu.SetFocus
    ElseIf txt_interno = "" Then
        MsgBox "Informe o nome interno.", vbInformation, "Atenção!"
        txt_interno.SetFocus
    ElseIf txt_disco = "" Then
        MsgBox "Informe o nome em disco.", vbInformation, "Atenção!"
        txt_disco.SetFocus
    ElseIf cbo_tipo.ListIndex = -1 Then
        MsgBox "Informe o tipo do programa.", vbInformation, "Atenção!"
        cbo_tipo.SetFocus
    ElseIf txt_configuravel <> "S" And txt_configuravel <> "N" Then
        MsgBox "Informe no quadro configurável 'Sim' ou 'não'.", vbInformation, "Atenção!"
        txt_configuravel.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_pesquisa_Click()
    consulta_programa.Show 1
    If Len(g_string) > 0 Then
        lNomeDisco = RetiraGString(1)
        If Programa.LocalizarNomeDisco(lNomeDisco) Then
            AtualTela
        End If
    End If
End Sub
Private Sub cmd_primeiro_Click()
    If Programa.LocalizarPrimeiro Then
        AtualTela
        cmd_proximo.SetFocus
    Else
        MsgBox "Não há registro.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    If Programa.LocalizarProximo Then
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
    If Programa.LocalizarUltimo Then
        AtualTela
        cmd_anterior.SetFocus
    Else
        MsgBox "Não há registro.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub Form_Activate()
    If lFlagCadastro = 0 Then
        DesativaBotoes
        If Programa.LocalizarUltimo Then
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
    PreencheCboTipo
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub PreencheCboTipo()
    cbo_tipo.Clear
    cbo_tipo.AddItem "CA"
    cbo_tipo.AddItem "CO"
    cbo_tipo.AddItem "ES"
    cbo_tipo.AddItem "GR"
    cbo_tipo.AddItem "MO"
    cbo_tipo.AddItem "RE"
End Sub

Private Sub txt_configuravel_GotFocus()
    If lOpcao = 1 And txt_configuravel = "" Then
        txt_configuravel = "S"
    End If
End Sub
Private Sub txt_configuravel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub txt_disco_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_observacao.SetFocus
    End If
End Sub
Private Sub txt_interno_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_disco.SetFocus
    End If
End Sub
Private Sub txt_menu_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_interno.SetFocus
    End If
End Sub
Private Sub txt_observacao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_tipo.SetFocus
    End If
End Sub
