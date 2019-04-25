VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form cadastro_dependente 
   Caption         =   "Cadastro de Dependente"
   ClientHeight    =   3375
   ClientLeft      =   120
   ClientTop       =   1875
   ClientWidth     =   7155
   Icon            =   "cad_dependente.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "cad_dependente.frx":030A
   ScaleHeight     =   3375
   ScaleWidth      =   7155
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   120
      Picture         =   "cad_dependente.frx":0750
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Cria um novo registro."
      Top             =   2400
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   1020
      Picture         =   "cad_dependente.frx":1DE2
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Altera o registro atual."
      Top             =   2400
      Width           =   795
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   1920
      Picture         =   "cad_dependente.frx":32DC
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Exclui o registro atual."
      Top             =   2400
      Width           =   795
   End
   Begin VB.CommandButton cmd_pesquisa 
      Caption         =   "&Pesquisa"
      Height          =   855
      Left            =   2820
      Picture         =   "cad_dependente.frx":496E
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Pesquisa um registro específico."
      Top             =   2400
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3720
      Picture         =   "cad_dependente.frx":5DE0
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   2400
      Width           =   795
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   6915
      Begin MSAdodcLib.Adodc adodc_funcionario 
         Height          =   330
         Left            =   3900
         Top             =   180
         Visible         =   0   'False
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "adodc_funcionario"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.TextBox txt_funcionario 
         Height          =   315
         Left            =   1860
         MaxLength       =   3
         TabIndex        =   2
         Top             =   180
         Width           =   375
      End
      Begin MSDataListLib.DataCombo dtcbo_funcionario 
         Bindings        =   "cad_dependente.frx":7472
         Height          =   315
         Left            =   2340
         TabIndex        =   3
         Top             =   180
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   "dtcbo_funcionario"
      End
      Begin VB.Label Label3 
         Caption         =   "&Funcionário"
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   1
         Top             =   180
         Width           =   1695
      End
   End
   Begin VB.Frame frm_dados 
      Enabled         =   0   'False
      Height          =   1755
      Left            =   120
      TabIndex        =   4
      Top             =   540
      Width           =   6915
      Begin VB.CheckBox chk_invalido 
         Caption         =   "&Inválido"
         Height          =   255
         Left            =   5100
         TabIndex        =   13
         Top             =   1380
         Width           =   1575
      End
      Begin VB.TextBox txt_codigo 
         Height          =   285
         Left            =   1860
         MaxLength       =   3
         TabIndex        =   6
         Top             =   240
         Width           =   435
      End
      Begin VB.TextBox txt_nome 
         Height          =   285
         Left            =   1860
         MaxLength       =   40
         TabIndex        =   8
         Top             =   600
         Width           =   4935
      End
      Begin MSMask.MaskEdBox msk_data_nascimento 
         Height          =   300
         Left            =   1860
         TabIndex        =   12
         Top             =   1320
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSAdodcLib.Adodc adodc_grau_dependencia 
         Height          =   330
         Left            =   3840
         Top             =   960
         Visible         =   0   'False
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "adodc_grau_dependencia"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataListLib.DataCombo dtcbo_grau_dependencia 
         Bindings        =   "cad_dependente.frx":7492
         Height          =   315
         Left            =   1860
         TabIndex        =   10
         Top             =   960
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   "dtcbo_grau_dependencia"
      End
      Begin VB.Label Label3 
         Caption         =   "Código do dependente"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "&Grau de dependência"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "&Nome de dependente"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "&Data de nascimento"
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   1695
      End
   End
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   4860
      TabIndex        =   21
      Top             =   2280
      Width           =   2175
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "cad_dependente.frx":74B7
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "cad_dependente.frx":8A39
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "cad_dependente.frx":9EAB
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "cad_dependente.frx":B3A5
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   6240
      Picture         =   "cad_dependente.frx":C89F
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Cancela o registro atual."
      Top             =   2400
      Width           =   795
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   5340
      Picture         =   "cad_dependente.frx":DD99
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Confirma o registro atual."
      Top             =   2400
      Width           =   795
   End
End
Attribute VB_Name = "cadastro_dependente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lFlagCadastro As Integer
Dim lOpcao As Integer
Dim lCodigoFuncionario As Integer
Dim lCodigoDependente As Integer
Dim lEmpresa As Integer
Private Funcionario As New cFuncionario
Private Dependente As New cDependente
Private Sub AtivaBotoes()
    cmd_novo.Enabled = True
    cmd_alterar.Enabled = True
    cmd_pesquisa.Enabled = True
    cmd_excluir.Enabled = True
    cmd_sair.Enabled = True
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
    frm_move.Visible = True
    txt_codigo.Enabled = True
    Frame1.Enabled = True
End Sub
Private Sub AtualTabe()
    Dependente.Empresa = g_empresa
    Dependente.CodigoFuncionario = Val(dtcbo_funcionario.BoundText)
    Dependente.CodigoDependente = Val(txt_codigo.Text)
    Dependente.Nome = txt_nome.Text
    Dependente.GrauDependencia = Val(dtcbo_grau_dependencia.BoundText)
    Dependente.DataNascimento = msk_data_nascimento
    Dependente.Invalido = chk_invalido.Value
End Sub
Private Sub AtualTela()
    lCodigoFuncionario = Dependente.CodigoFuncionario
    lCodigoDependente = Dependente.CodigoDependente
    txt_funcionario.Text = Dependente.CodigoFuncionario
    dtcbo_funcionario.BoundText = Dependente.CodigoFuncionario
    txt_codigo.Text = Dependente.CodigoDependente
    txt_nome.Text = Dependente.Nome
    dtcbo_grau_dependencia.BoundText = Dependente.GrauDependencia
    msk_data_nascimento = Dependente.DataNascimento
    If Dependente.Invalido Then
        chk_invalido.Value = 1
    Else
        chk_invalido.Value = 0
    End If
    frm_dados.Enabled = False
End Sub
'Function BuscaDados() As Boolean
'    BuscaDados = False
'    With tbl_dependente
'        If .RecordCount > 0 Then
'            .Seek "<", g_empresa, 9999, 9999
'            If Not .NoMatch Then
'                If !Empresa = g_empresa Then
'                    AtualTela
'                    BuscaDados = True
'                End If
'            End If
'        End If
'    End With
'End Function
Private Sub CriaNovoDependente()
    txt_codigo.Text = 1
    If Dependente.LocalizarUltimoDependente(g_empresa, Val(txt_funcionario.Text)) Then
        txt_codigo.Text = Dependente.CodigoDependente + 1
    End If
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
    Set Dependente = Nothing
    Set Funcionario = Nothing
    frm_cadastro.Show
End Sub
Private Sub Inclui()
    lOpcao = 1
    DesativaBotoes
    cmd_novo.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
End Sub
Private Sub chk_invalido_KeyPress(KeyAscii As Integer)
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
    Frame1.Enabled = False
    txt_codigo.Enabled = False
    txt_nome.SetFocus
End Sub
Private Sub cmd_anterior_Click()
    If Dependente.LocalizarAnterior() Then
        AtualTela
    Else
        MsgBox "Início de Arquivo.", vbInformation, "Atenção!"
        cmd_proximo.SetFocus
    End If
End Sub
Private Sub cmd_cancelar_Click()
    LimpaTela
    LimpaTela2
    If Dependente.LocalizarCodigo(g_empresa, lCodigoFuncionario, lCodigoDependente) Then
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
    txt_funcionario = ""
    dtcbo_funcionario.BoundText = ""
End Sub
Private Sub LimpaTela2()
    txt_codigo = ""
    txt_nome = ""
    dtcbo_grau_dependencia.BoundText = ""
    msk_data_nascimento = "__/__/____"
    chk_invalido.Value = 0
End Sub
Private Sub cmd_excluir_Click()
    If Val(txt_codigo.Text) > 0 Then
        If (MsgBox("Deseja Realmente Excluir Este Registro?", 4 + 32 + 256, "Exclusão de Registro!")) = 6 Then
            If Dependente.Excluir(g_empresa, lCodigoFuncionario, lCodigoDependente) Then
                LimpaTela
                LimpaTela2
                If Not Dependente.LocalizarUltimoDependente(g_empresa, lCodigoFuncionario) Then
                    If Not Dependente.LocalizarUltimo(g_empresa) Then
                        DesativaBotoes
                        cmd_novo.Enabled = True
                        cmd_sair.Enabled = True
                        cmd_novo.SetFocus
                        Exit Sub
                    End If
                End If
                AtualTela
            Else
                MsgBox "Não foi possivel excluir este registro!", vbInformation, "Erro de Verificação!"
            End If
        End If
    End If
End Sub
Private Sub cmd_novo_Click()
    LimpaTela
    LimpaTela2
    Inclui
    frm_dados.Enabled = True
    txt_funcionario.SetFocus
End Sub
Private Sub cmd_ok_Click()
    On Error GoTo FileError
    If ValidaCampos Then
        AtivaBotoes
        If lOpcao = 1 Then
            AtualTabe
            If Dependente.Incluir Then
                lCodigoFuncionario = Val(txt_funcionario.Text)
                lCodigoDependente = Val(txt_codigo.Text)
            Else
                MsgBox "Não foi possível incluir este registro!", vbInformation, "Erro de Verificação!"
            End If
        ElseIf lOpcao = 2 Then
            AtualTabe
            If Not Dependente.Alterar(g_empresa, lCodigoFuncionario, lCodigoDependente) Then
                MsgBox "Não foi possível alterar este registro!", vbInformation, "Erro de Verificação!"
            End If
        End If
        lOpcao = 0
        If Dependente.LocalizarCodigo(g_empresa, lCodigoFuncionario, lCodigoDependente) Then
            AtualTela
        End If
        cmd_novo.SetFocus
    End If
    Exit Sub
FileError:
    MsgBox Error
    'ErroArquivo tbl_dependente.Name, "Dependenteo"
    Exit Sub
End Sub
'Function BuscaRegistro(x_codigo_funcionario As Integer, x_codigo_dependente As Integer) As Boolean
'    BuscaRegistro = False
'    tbl_dependente.Seek "=", g_empresa, x_codigo_funcionario, x_codigo_dependente
'    If Not tbl_dependente.NoMatch Then
'        BuscaRegistro = True
'        AtualTela
'    End If
'End Function
Function ValidaCampos() As Integer
    ValidaCampos = False
    If Val(dtcbo_funcionario.BoundText) = 0 Then
        MsgBox "Escolha um funcionário.", vbInformation, "Atenção!"
        dtcbo_funcionario.SetFocus
    ElseIf txt_codigo = "" Then
        MsgBox "Informe o código do dependente.", vbInformation, "Atenção!"
        txt_codigo.SetFocus
    ElseIf txt_nome = "" Then
        MsgBox "Informe o nome do dependente.", vbInformation, "Atenção!"
        txt_nome.SetFocus
    ElseIf Val(dtcbo_grau_dependencia.BoundText) = 0 Then
        MsgBox "Escolha o grau de dependência.", vbInformation, "Atenção!"
        dtcbo_grau_dependencia.SetFocus
    ElseIf Not IsDate(msk_data_nascimento) Then
        MsgBox "Informe a data de nascimento.", vbInformation, "Atenção!"
        msk_data_nascimento.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_pesquisa_Click()
    consulta_dependente.Show 1
    If Len(g_string) > 0 Then
        lCodigoFuncionario = RetiraGString(1)
        lCodigoDependente = RetiraGString(2)
        If Dependente.LocalizarCodigo(g_empresa, lCodigoFuncionario, lCodigoDependente) Then
            AtivaBotoes
            AtualTela
        End If
    End If
End Sub
Private Sub cmd_primeiro_Click()
    If Dependente.LocalizarPrimeiro() Then
        AtualTela
        cmd_proximo.SetFocus
    Else
        MsgBox "Não há registro.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    If Dependente.LocalizarProximo() Then
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
    If Dependente.LocalizarUltimoDependente(g_empresa, Val(txt_funcionario.Text)) Then
        AtualTela
        AtivaBotoes
        cmd_anterior.SetFocus
    Else
        MsgBox "Não há dependentes para este funcionário.", vbInformation, "Registro não encontrado!"
    End If
End Sub
Private Sub dtcbo_funcionario_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        If lOpcao <> 0 Then
            txt_nome.SetFocus
        Else
            LimpaTela2
            If cmd_ultimo.Visible Then
                cmd_ultimo.SetFocus
            Else
                cmd_novo.SetFocus
            End If
        End If
    End If
End Sub
Private Sub dtcbo_funcionario_LostFocus()
    If dtcbo_funcionario.BoundText <> "" And lOpcao > 0 Then
        txt_funcionario = dtcbo_funcionario.BoundText
        txt_funcionario_LostFocus
        txt_nome.SetFocus
    End If
    If lOpcao = 0 Then
        txt_funcionario = dtcbo_funcionario.BoundText
        DesativaBotoes
        cmd_novo.Enabled = True
        cmd_pesquisa.Enabled = True
        cmd_sair.Enabled = True
        cmd_ultimo_Click
    End If
End Sub
Private Sub dtcbo_grau_dependencia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data_nascimento.SetFocus
    End If
End Sub
Private Sub dtcbo_grau_dependencia_LostFocus()
    'If lOpcao = 1 And Val(dtcbo_grau_dependencia.BoundText) > 0 Then
    '    lCodigoDependente = 0
    '    tbl_dependente.Index = "id_codigo"
    '    tbl_dependente.Seek "<", Val(dtcbo_grau_dependencia.BoundText), 9999
    '    If Not tbl_dependente.NoMatch Then
    '        If tbl_dependente![Codigo do Grupo] = Val(dtcbo_grau_dependencia.BoundText) Then
    '            lCodigoDependente = tbl_dependente![Codigo do Produto]
    '        End If
    '    End If
    '    lCodigoDependente = lCodigoDependente + 1
    '    tbl_dependente.Index = "id_codigo"
    'End If
End Sub
Private Sub Form_Activate()
    If g_empresa <> lEmpresa Then
        lFlagCadastro = 0
    End If
    If lFlagCadastro = 0 Then
        lOpcao = 0
        lEmpresa = g_empresa
        adodc_funcionario.ConnectionString = gConnectionString
        adodc_funcionario.RecordSource = "SELECT Codigo, Nome FROM Funcionario WHERE Empresa = " & g_empresa & " ORDER BY Nome"
        adodc_funcionario.Refresh
        DesativaBotoes
        If Dependente.LocalizarUltimo(g_empresa) Then
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
    adodc_grau_dependencia.ConnectionString = gConnectionString
    adodc_grau_dependencia.RecordSource = "SELECT Codigo, Nome FROM Grau_Dependencia ORDER BY Nome"
    adodc_grau_dependencia.Refresh
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_nascimento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        chk_invalido.SetFocus
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
    If lOpcao = 1 And Val(txt_codigo.Text) > 0 Then
        If Dependente.LocalizarCodigo(g_empresa, Val(txt_funcionario.Text), Val(txt_codigo.Text)) Then
            MsgBox "Já existe dependente cadastrado com este código." & Chr(10) & Chr(10) & "Mude o código informado.", vbInformation, "Duplicidade de Registro!"
            txt_codigo.SetFocus
        End If
    End If
End Sub
Private Sub txt_funcionario_GotFocus()
    txt_funcionario.SelStart = 0
    txt_funcionario.SelLength = Len(txt_funcionario.Text)
End Sub
Private Sub txt_funcionario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dtcbo_funcionario.SetFocus
    End If
End Sub
Private Sub txt_funcionario_LostFocus()
    If Val(txt_funcionario.Text) > 0 Then
        If Funcionario.LocalizarCodigo(g_empresa, Val(txt_funcionario.Text)) Then
            If Funcionario.Situacao = "I" Then
                MsgBox "O funcionário " & Trim(Funcionario.Nome) & " está inativo.", vbInformation, "Atenção!"
                txt_funcionario.SetFocus
                Exit Sub
            Else
                dtcbo_funcionario.BoundText = Val(txt_funcionario.Text)
                If lOpcao = 1 Then
                    CriaNovoDependente
                    txt_nome.SetFocus
                ElseIf lOpcao = 2 Then
                    txt_nome.SetFocus
                End If
            End If
        Else
            MsgBox "Funcionário não cadastrado.", vbInformation, "Atenção!"
            txt_funcionario.SetFocus
            Exit Sub
        End If
    End If
End Sub
Private Sub txt_nome_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dtcbo_grau_dependencia.SetFocus
    End If
End Sub
Private Sub txt_nome_LostFocus()
    If lOpcao = 1 And txt_nome.Text <> "" And Val(txt_funcionario.Text) > 0 Then
        If Dependente.LocalizarNome(g_empresa, Val(txt_funcionario.Text), txt_nome.Text) Then
            If (MsgBox("Já existe Dependente cadastrado com este nome." & Chr(10) & Chr(10) & "Deseja cadastrar assim mesmo?", 4 + 32 + 256, "Duplicidade de Registro!")) = 7 Then
                txt_nome.SetFocus
            End If
        End If
    End If
End Sub
