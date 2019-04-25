VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C21-0000-0000-0000-000000000046}#4.0#0"; "TDBG32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form movimento_falta_funcionario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimento de Faltas de Funcionários"
   ClientHeight    =   5655
   ClientLeft      =   2610
   ClientTop       =   1905
   ClientWidth     =   6930
   Icon            =   "mov_falta_funcionario.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "mov_falta_funcionario.frx":030A
   ScaleHeight     =   5655
   ScaleWidth      =   6930
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   120
      Picture         =   "mov_falta_funcionario.frx":0750
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Cria um novo registro."
      Top             =   4680
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   1020
      Picture         =   "mov_falta_funcionario.frx":1DE2
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Altera o registro atual."
      Top             =   4680
      Width           =   795
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   1920
      Picture         =   "mov_falta_funcionario.frx":32DC
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Exclui o registro atual."
      Top             =   4680
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Caption         =   "&Sair"
      Height          =   855
      Left            =   2820
      Picture         =   "mov_falta_funcionario.frx":496E
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   4680
      Width           =   795
   End
   Begin VB.Data dta_movimento_falta_funcionario 
      Caption         =   "dta_movimento_falta_funcionario"
      Connect         =   "Access"
      DatabaseName    =   "Sgp_data.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Movimento_Falta_Funcionario"
      Top             =   4200
      Visible         =   0   'False
      Width           =   3675
   End
   Begin VB.Frame frmDados 
      Enabled         =   0   'False
      Height          =   2595
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6675
      Begin VB.CheckBox chkjustificada 
         Caption         =   "Justificada"
         Height          =   195
         Left            =   3300
         TabIndex        =   10
         Top             =   1800
         Width           =   1395
      End
      Begin VB.Data dta_funcionario 
         Caption         =   "dta_funcionario"
         Connect         =   "Access"
         DatabaseName    =   "Sgp_data.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   3720
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   660
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox txt_motivo 
         Height          =   315
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   12
         Top             =   2100
         Width           =   4815
      End
      Begin VB.Frame frmTipoFalta 
         Caption         =   "Tipo de Falta"
         Height          =   555
         Left            =   180
         TabIndex        =   5
         Top             =   1080
         Width           =   3555
         Begin VB.OptionButton optTipoFalta 
            Caption         =   "Período &Integral"
            Height          =   255
            Index           =   1
            Left            =   1860
            TabIndex        =   7
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton optTipoFalta 
            Caption         =   "&Meio Período"
            Height          =   255
            Index           =   0
            Left            =   300
            TabIndex        =   6
            Top             =   240
            Width           =   1455
         End
      End
      Begin MSDBCtls.DBCombo dbcbo_funcionario 
         Bindings        =   "mov_falta_funcionario.frx":6000
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
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.CheckBox chkAbonada 
         Caption         =   "Abonada"
         Height          =   195
         Left            =   1680
         TabIndex        =   9
         Top             =   1800
         Width           =   1395
      End
      Begin VB.Label Label1 
         Caption         =   "Situação da falta"
         Height          =   255
         Index           =   3
         Left            =   180
         TabIndex        =   8
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Motivo"
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   11
         Top             =   2160
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
   Begin TrueDBGrid.TDBGrid TDBGrid1 
      Bindings        =   "mov_falta_funcionario.frx":601E
      Height          =   1815
      Left            =   120
      OleObjectBlob   =   "mov_falta_funcionario.frx":604C
      TabIndex        =   13
      Top             =   2760
      Width           =   6675
   End
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   4620
      TabIndex        =   20
      Top             =   4560
      Width           =   2175
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "mov_falta_funcionario.frx":7C3C
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "mov_falta_funcionario.frx":9136
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "mov_falta_funcionario.frx":A630
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "mov_falta_funcionario.frx":BAA2
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   5100
      Picture         =   "mov_falta_funcionario.frx":D024
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Confirma o registro atual."
      Top             =   4680
      Width           =   795
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   6000
      Picture         =   "mov_falta_funcionario.frx":E62E
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Cancela o registro atual."
      Top             =   4680
      Width           =   795
   End
End
Attribute VB_Name = "movimento_falta_funcionario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lOpcao As Integer
Dim l_data As Date
Dim l_funcionario As Integer
Dim l_campo As String
Dim l_arquivo As String
Dim l_condicao As String
Dim l_ordem As String
Dim l_sql As String
Dim l_movimento_falta_funcionario As Integer
Dim tbl_funcionario As Table
Dim tbl_movimento_falta_funcionario As Table
Private Sub AtualizaGrid()
    'início das variáveis para traduzir
    Dim VItem As New ValueItem
    Dim VItems As ValueItems
    'fim das variáveis para traduzir
    Dim i As Integer
    l_campo = "Select Movimento_Falta_Funcionario.Data, Movimento_Falta_Funcionario.[Codigo do Funcionario], Movimento_Falta_Funcionario.[Tipo de Falta], Movimento_Falta_Funcionario.Motivo"
    l_arquivo = " From Movimento_Falta_Funcionario"
    l_condicao = " Where "
    l_condicao = l_condicao & " Movimento_Falta_Funcionario.Empresa = " & g_empresa
    l_ordem = " order by Movimento_Falta_Funcionario.Data"
    l_sql = l_campo & l_arquivo & l_condicao & l_ordem
    dta_movimento_falta_funcionario.RecordSource = l_sql
    dta_movimento_falta_funcionario.Refresh
    While TDBGrid1.Columns.Count <> 0
        TDBGrid1.Columns.Remove 0
    Wend
    For i = 0 To 3
        TDBGrid1.Columns.Add 0
    Next
    For i = 0 To 3
        TDBGrid1.Columns(i).Visible = True
    Next
'    data1.RecordSource = "Select Data, nome, [Tipo de Falta], Motivo from Movimento_Falta_Funcionario, Funcionarios, Movimento_Falta_Funcionario inner join Funcionarios on Movimento_Falta_Funcionario.[Codigo do Funcionario] = Funcionarios.Codigo and Funcionarios.Empresa = " & g_empresa
'    data1.Refresh
    TDBGrid1.Columns(0).DataField = "Data"
    TDBGrid1.Columns(0).Caption = "Data"
    TDBGrid1.Columns(0).Alignment = dbgCenter
    TDBGrid1.Columns(0).Width = 1000
    TDBGrid1.Columns(1).DataField = "Codigo do Funcionario"
    TDBGrid1.Columns(1).Caption = "Nome do Funcionário"
    TDBGrid1.Columns(1).Alignment = dbgLeft
    TDBGrid1.Columns(1).Width = 3000
    TDBGrid1.Columns(2).DataField = "Tipo de Falta"
    TDBGrid1.Columns(2).Caption = "Tipo de Falta"
    TDBGrid1.Columns(2).Alignment = dbgLeft
    TDBGrid1.Columns(2).Width = 1500
    TDBGrid1.Columns(3).DataField = "Motivo"
    TDBGrid1.Columns(3).Caption = "Motivo da Falta"
    TDBGrid1.Columns(3).Alignment = dbgLeft
    TDBGrid1.Columns(3).Width = 2500
    'Início da Tradução
    Set VItems = TDBGrid1.Columns(2).ValueItems
    VItem.Value = 1
    VItem.DisplayValue = "Meio Período"
    VItems.Add VItem
    VItem.Value = 2
    VItem.DisplayValue = "Período Integral"
    VItems.Add VItem
    VItems.Translate = True
    'Traduz Funcionário
    Set VItems = TDBGrid1.Columns(1).ValueItems
    With tbl_funcionario
        .Seek ">=", g_empresa, 0
        Do Until .EOF
            If !Empresa <> g_empresa Then
                Exit Do
            End If
            VItem.Value = !Codigo
            VItem.DisplayValue = !Nome
            VItems.Add VItem
            .MoveNext
        Loop
    End With
    VItems.Translate = True
End Sub
Private Sub AtualTabe()
    With tbl_movimento_falta_funcionario
        l_funcionario = Val(dbcbo_funcionario.BoundText)
        l_data = msk_data.Text
        !Empresa = g_empresa
        !Data = msk_data.Text
        ![Codigo do Funcionario] = Val(dbcbo_funcionario.BoundText)
        If optTipoFalta(0) Then
            ![Tipo de Falta] = 1
        Else
            ![Tipo de Falta] = 2
        End If
        !Abonada = chkAbonada.Value
        !Justificada = chkjustificada.Value
        !Motivo = txt_motivo.Text
    End With
End Sub
Private Sub AtualTela()
    With tbl_movimento_falta_funcionario
        l_funcionario = ![Codigo do Funcionario]
        l_data = !Data
        
        msk_data.Text = Format(!Data, "dd/mm/yyyy")
        dbcbo_funcionario.BoundText = ![Codigo do Funcionario]
        If ![Tipo de Falta] = 1 Then
            optTipoFalta(0) = True
        Else
            optTipoFalta(1) = True
        End If
        chkAbonada.Value = !Abonada
        chkjustificada.Value = !Justificada
        txt_motivo.Text = !Motivo
        frmDados.Enabled = False
    End With
End Sub
Function ExisteDados() As Boolean
    ExisteDados = False
    With tbl_movimento_falta_funcionario
        If .RecordCount > 0 Then
            .Seek "<", g_empresa, CDate("31/12/2500"), 9999
            If Not .NoMatch Then
                If !Empresa = g_empresa Then
                    ExisteDados = True
                End If
            End If
        End If
    End With
End Function
Private Sub Finaliza()
    tbl_funcionario.Close
    tbl_movimento_falta_funcionario.Close
End Sub
Private Sub Incluir()
    lOpcao = 1
    DesativaBotoes
    cmd_novo.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
End Sub
Function BuscaRegistro(x_data As Date, x_funcionario As Integer) As Boolean
    BuscaRegistro = False
    With tbl_movimento_falta_funcionario
        .Seek "=", g_empresa, x_data, x_funcionario
        If Not .NoMatch Then
            AtualTela
            BuscaRegistro = True
        End If
    End With
End Function
Private Sub TabelaFuncionarioRefresh()
    dta_funcionario.RecordSource = "Select * From Funcionario Where Empresa = " & g_empresa & " And Situacao = 'A' Order By Nome"
    dta_funcionario.Refresh
End Sub
Private Sub chkAbonada_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_motivo.SetFocus
    End If
End Sub
Private Sub chkJustificada_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_motivo.SetFocus
    End If
End Sub
Private Sub cmd_alterar_Click()
    lOpcao = 2
    DesativaBotoes
    cmd_alterar.Visible = True
    cmd_alterar.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    frmDados.Enabled = True
    If optTipoFalta(0) Then
        optTipoFalta(0).SetFocus
    Else
        optTipoFalta(1).SetFocus
    End If
End Sub
Private Sub cmd_anterior_Click()
    Dim x_flag As Integer
    If tbl_movimento_falta_funcionario.RecordCount > 0 Then
        tbl_movimento_falta_funcionario.MovePrevious
        If tbl_movimento_falta_funcionario.BOF Then
            x_flag = 1
        Else
            If tbl_movimento_falta_funcionario!Empresa <> g_empresa Then
                x_flag = 1
            End If
        End If
        If x_flag = 1 Then
            MsgBox "Início de Arquivo.", 48, "Atenção!"
            tbl_movimento_falta_funcionario.MoveNext
            cmd_proximo.SetFocus
        Else
            AtualTela
        End If
    End If
End Sub
Private Sub cmd_cancelar_Click()
    LimpaTela
    If BuscaRegistro(l_data, l_funcionario) Then
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
    dbcbo_funcionario.BoundText = 0
    optTipoFalta(0) = True
    chkAbonada.Value = False
    chkjustificada.Value = False
    txt_motivo.Text = ""
End Sub
Private Sub cmd_cancelar_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    cmd_cancelar.Value = True
End Sub
Private Sub cmd_excluir_Click()
    If IsDate(msk_data.Text) Then
        If (MsgBox("Deseja Realmente Excluir Este Registro?", 4 + 32 + 256, "Exclusão de Registro!")) = 6 Then
            tbl_movimento_falta_funcionario.Edit
            tbl_movimento_falta_funcionario.Delete
            LimpaTela
            AtualizaGrid
            If Not BuscaDados Then
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
            tbl_movimento_falta_funcionario.AddNew
            AtualTabe
            tbl_movimento_falta_funcionario.Update
        ElseIf lOpcao = 2 Then
            tbl_movimento_falta_funcionario.Edit
            AtualTabe
            tbl_movimento_falta_funcionario.Update
        End If
        AtualizaGrid
        If BuscaRegistro(l_data, l_funcionario) Then
            lOpcao = 0
            cmd_novo.SetFocus
        End If
    End If
    Exit Sub
FileError:
    ErroArquivo tbl_movimento_falta_funcionario.name, "Movimentoo"
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
    ElseIf optTipoFalta(0) = False And optTipoFalta(1) = False Then
        MsgBox "Escolha o tipo de falta.", 64, "Atenção!"
        optTipoFalta(0).SetFocus
    ElseIf txt_motivo = "" Then
        MsgBox "Informe o motivo.", 64, "Atenção!"
        txt_motivo.SetFocus
    ElseIf chkjustificada.Value = True And chkAbonada.Value = True Then
        MsgBox "A situação da falta não pode ser abonada e justificada.", 64, "Atenção!"
        chkAbonada.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_primeiro_Click()
    Dim x_flag As Integer
    If tbl_movimento_falta_funcionario.RecordCount > 0 Then
        tbl_movimento_falta_funcionario.Seek ">", g_empresa, CDate("01/01/1900"), 0
        If Not tbl_movimento_falta_funcionario.NoMatch Then
            If tbl_movimento_falta_funcionario!Empresa <> g_empresa Then
                x_flag = 1
            End If
        Else
            x_flag = 1
        End If
        If x_flag = 1 Then
            MsgBox "Não há registros nesta empresa.", 64, "Erro de Verificação!"
        Else
            AtualTela
            cmd_proximo.SetFocus
        End If
    End If
End Sub
Private Sub cmd_proximo_Click()
    Dim x_flag As Integer
    If tbl_movimento_falta_funcionario.RecordCount > 0 Then
        tbl_movimento_falta_funcionario.MoveNext
        If tbl_movimento_falta_funcionario.EOF Then
            x_flag = 1
        Else
            If tbl_movimento_falta_funcionario!Empresa <> g_empresa Then
                x_flag = 1
            End If
        End If
        If x_flag = 1 Then
            MsgBox "Fim de Arquivo.", 48, "Atenção!"
            tbl_movimento_falta_funcionario.MovePrevious
            cmd_anterior.SetFocus
        Else
            AtualTela
        End If
    End If
End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub cmd_ultimo_Click()
    Dim x_flag As Integer
    If tbl_movimento_falta_funcionario.RecordCount > 0 Then
        tbl_movimento_falta_funcionario.Seek "<", g_empresa, CDate("31/12/2500"), 9999
        If Not tbl_movimento_falta_funcionario.NoMatch Then
            If tbl_movimento_falta_funcionario!Empresa <> g_empresa Then
                x_flag = 1
            End If
        Else
            x_flag = 1
        End If
        If x_flag = 1 Then
            MsgBox "Não há registros nesta empresa.", 64, "Erro de Verificação!"
        Else
            AtualTela
            cmd_anterior.SetFocus
        End If
    End If
End Sub
Private Sub dbcbo_funcionario_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys "{Tab}"
    End If
End Sub
Private Sub dbcbo_funcionario_LostFocus()
    If Val(dbcbo_funcionario.BoundText) > 0 Then
        tbl_movimento_falta_funcionario.Seek "=", g_empresa, msk_data, Val(dbcbo_funcionario.BoundText)
        If Not tbl_movimento_falta_funcionario.NoMatch Then
            MsgBox "Registro jâ Cadastrado", vbExclamation, "Atenção!"
            dbcbo_funcionario.SetFocus
        End If
    End If
End Sub
Private Sub Form_Activate()
    TabelaFuncionarioRefresh
    If l_movimento_falta_funcionario = 0 Then
        DesativaBotoes
        If BuscaDados Then
            AtivaBotoes
        Else
            cmd_novo.Enabled = True
            cmd_sair.Enabled = True
        End If
        cmd_novo.SetFocus
    Else
        l_movimento_falta_funcionario = 0
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
    TDBGrid1.Enabled = True
    dta_movimento_falta_funcionario.Enabled = True
End Sub
Private Sub DesativaBotoes()
    TDBGrid1.Enabled = False
    dta_movimento_falta_funcionario.Enabled = False
    cmd_novo.Enabled = False
    cmd_excluir.Enabled = False
    cmd_alterar.Enabled = False
    cmd_sair.Enabled = False
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
    frm_move.Visible = False
End Sub
Function BuscaDados() As Boolean
    BuscaDados = False
    With tbl_movimento_falta_funcionario
        If .RecordCount > 0 Then
            .Seek "<", g_empresa, CDate("31/12/2500"), 9999
            If Not .NoMatch Then
                If !Empresa = g_empresa Then
                    AtualTela
                    BuscaDados = True
                End If
            End If
        End If
    End With
End Function
Private Sub Form_Deactivate()
    l_movimento_falta_funcionario = 1
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
    Set tbl_funcionario = bd_sgp.OpenTable("Funcionario")
    Set tbl_movimento_falta_funcionario = bd_sgp.OpenRecordset("Movimento_Falta_Funcionario", dbOpenTable)
    tbl_funcionario.Index = "id_codigo"
    tbl_movimento_falta_funcionario.Index = "id_data"
    AtualizaGrid
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{Tab}"
    End If
End Sub
Private Sub optTipoFalta_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{Tab}"
    End If
End Sub
Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If ExisteDados And lOpcao <> 1 Then
'        tbl_movimento_falta_funcionario.Seek "=", tbl_movimento_falta_funcionario!codigo
'        If Not tbl_movimento_falta_funcionario.NoMatch Then
            AtualTela
'        End If
    End If
End Sub
Private Sub txt_motivo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{Tab}"
    End If
End Sub
