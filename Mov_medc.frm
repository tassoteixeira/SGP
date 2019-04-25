VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form mov_medicao_combustiveis 
   Caption         =   "Medição dos Combustíveis"
   ClientHeight    =   5670
   ClientLeft      =   2040
   ClientTop       =   1875
   ClientWidth     =   6885
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "Mov_medc.frx":0000
   ScaleHeight     =   5670
   ScaleWidth      =   6885
   Begin VB.CommandButton cmd_pesquisa 
      Caption         =   "&Pesquisa"
      Height          =   855
      Left            =   2760
      Picture         =   "Mov_medc.frx":0446
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Pesquisa um registro específico."
      Top             =   4740
      Width           =   795
   End
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   60
      Picture         =   "Mov_medc.frx":1720
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Cria um novo registro."
      Top             =   4740
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   960
      Picture         =   "Mov_medc.frx":29FA
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Altera o registro atual."
      Top             =   4740
      Width           =   795
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   1860
      Picture         =   "Mov_medc.frx":3CD4
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Exclui o registro atual."
      Top             =   4740
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3660
      Picture         =   "Mov_medc.frx":4FAE
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   4740
      Width           =   795
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2880
      TabIndex        =   35
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame frm_dados 
      Enabled         =   0   'False
      Height          =   4575
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6735
      Begin VB.CommandButton cmd_transfere_dados_lmc 
         Caption         =   "&Transfere p/ LMC"
         Height          =   675
         Left            =   5160
         Picture         =   "Mov_medc.frx":6288
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Transfere as entradas de combustíveis para o LMC."
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txt_observacao_3 
         Height          =   285
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   22
         Top             =   4080
         Width           =   4935
      End
      Begin VB.TextBox txt_observacao_2 
         Height          =   285
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   20
         Top             =   3720
         Width           =   4935
      End
      Begin VB.TextBox txt_observacao_1 
         Height          =   285
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   18
         Top             =   3360
         Width           =   4935
      End
      Begin VB.TextBox msk_tanque_6 
         Height          =   285
         Left            =   1680
         TabIndex        =   16
         Top             =   3000
         Width           =   1095
      End
      Begin VB.TextBox msk_tanque_5 
         Height          =   285
         Left            =   1680
         TabIndex        =   14
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox msk_tanque_4 
         Height          =   285
         Left            =   1680
         TabIndex        =   12
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox msk_tanque_3 
         Height          =   285
         Left            =   1680
         TabIndex        =   10
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox msk_tanque_2 
         Height          =   285
         Left            =   1680
         TabIndex        =   8
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox msk_tanque_1 
         Height          =   285
         Left            =   1680
         TabIndex        =   6
         Top             =   1200
         Width           =   1095
      End
      Begin VB.ComboBox cbo_combustivel 
         Height          =   300
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   720
         Width           =   4935
      End
      Begin MSMask.MaskEdBox msk_data 
         Height          =   300
         Left            =   1680
         TabIndex        =   2
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label11 
         Caption         =   "Observação Linha 3"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   4080
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Observação Linha 2"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Observação Linha 1"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Estoque Tanque 6"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Estoque Tanque 5"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Estoque Tanque 4"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Estoque Tanque 3"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Estoque Tanque 2"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Estoque Tanque 1"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Data da Medição"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Combustível"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   4620
      TabIndex        =   30
      Top             =   4620
      Width           =   2175
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "Mov_medc.frx":767A
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "Mov_medc.frx":8954
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "Mov_medc.frx":9C2E
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "Mov_medc.frx":AF08
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   5100
      Picture         =   "Mov_medc.frx":C1E2
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Confirma o registro atual."
      Top             =   4740
      Width           =   795
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   6000
      Picture         =   "Mov_medc.frx":D4BC
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Cancela o registro atual."
      Top             =   4740
      Width           =   795
   End
End
Attribute VB_Name = "mov_medicao_combustiveis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lOpcao As Integer
Dim lTipoCombustivel As String * 2
Dim lData As Date
Dim tbl_med_comb
Private Sub AtivaBotoes()
    cmd_novo.Enabled = True
    cmd_excluir.Enabled = True
    cmd_alterar.Enabled = True
    cmd_sair.Enabled = True
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
    cmd_transfere_dados_lmc.Visible = False
    frm_move.Visible = True
End Sub
Private Sub AtualTabe()
    lData = msk_data
    With tbl_med_comb
        !Empresa = g_empresa
        !Data = msk_data
        !tipo_combustivel = Mid(cbo_combustivel, 1, 2)
        !tanque_1 = fValidaValor1(msk_tanque_1)
        !tanque_2 = fValidaValor1(msk_tanque_2)
        !tanque_3 = fValidaValor1(msk_tanque_3)
        !tanque_4 = fValidaValor1(msk_tanque_4)
        !tanque_5 = fValidaValor1(msk_tanque_5)
        !tanque_6 = fValidaValor1(msk_tanque_6)
        !observacao_1 = Format(txt_observacao_1, "")
        !observacao_2 = Format(txt_observacao_2, "")
        !observacao_3 = Format(txt_observacao_3, "")
    End With
End Sub
Private Sub AtualTela()
    Dim i As Integer
    With tbl_med_comb
        lData = !Data
        lTipoCombustivel = !tipo_combustivel
        msk_data = Format(!Data, "dd/mm/yyyy")
        tbl_combustivel.Seek "=", g_empresa, !tipo_combustivel
        If Not tbl_combustivel.NoMatch Then
            For i = 0 To cbo_combustivel.ListCount - 1
                cbo_combustivel.ListIndex = i
                If Mid(cbo_combustivel, 1, 2) = tbl_combustivel!Codigo Then
                    Exit For
                End If
            Next
        Else
            cbo_combustivel.ListIndex = -1
        End If
        msk_tanque_1 = Format(!tanque_1, "###,##0.0")
        msk_tanque_2 = Format(!tanque_2, "###,##0.0")
        msk_tanque_3 = Format(!tanque_3, "###,##0.0")
        msk_tanque_4 = Format(!tanque_4, "###,##0.0")
        msk_tanque_5 = Format(!tanque_5, "###,##0.0")
        msk_tanque_6 = Format(!tanque_6, "###,##0.0")
        txt_observacao_1 = Format(!observacao_1, "")
        txt_observacao_2 = Format(!observacao_2, "")
        txt_observacao_3 = Format(!observacao_3, "")
    End With
    frm_dados.Enabled = True
End Sub
Function BuscaDados() As Boolean
    BuscaDados = False
    With tbl_med_comb
        If .RecordCount > 0 Then
            tbl_med_comb.Seek "<", g_empresa, CDate("31/12/2500"), "ZZ"
            If Not tbl_med_comb.NoMatch Then
                If tbl_med_comb!Empresa = g_empresa Then
                    BuscaDados = True
                    AtualTela
                    Exit Function
                End If
            End If
        End If
    End With
    LimpaTela
End Function
Function BuscaRegistro(x_data As Date, x_tipo_combustivel As String) As Boolean
    BuscaRegistro = False
    tbl_med_comb.Seek "=", g_empresa, x_data, x_tipo_combustivel
    If Not tbl_med_comb.NoMatch Then
        BuscaRegistro = True
        AtualTela
    End If
End Function
Private Sub DesativaBotoes()
    cmd_novo.Enabled = False
    cmd_excluir.Enabled = False
    cmd_alterar.Enabled = False
    cmd_sair.Enabled = False
    frm_move.Visible = False
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
    If g_nome_usuario = "L.M.C." Then
        cmd_transfere_dados_lmc.Visible = True
    End If
End Sub
Private Sub GambiDistribuiDiferenca()
    Dim quantidade_x As Long
    Dim data_x As Date
    data_x = "03/05/96"
    Do Until data_x = "31/05/96"
        data_x = data_x + 1
        tbl_med_comb.Seek "=", g_empresa, data_x, "D"
        If Not tbl_med_comb.NoMatch Then
            tbl_med_comb.Edit
            If tbl_med_comb!Data < "18/05/96" Then
                quantidade_x = quantidade_x + 2
            Else
                quantidade_x = quantidade_x + 2
            End If
            tbl_med_comb!tanque_1 = tbl_med_comb!tanque_1 - quantidade_x
            tbl_med_comb.Update
        End If
    Loop
End Sub
Private Sub GambiDeletaMedicao()
    tbl_med_comb.Seek ">=", g_empresa, CDate("02/03/97"), "D"
    Do Until tbl_med_comb.EOF
        If Trim(tbl_med_comb!tipo_combustivel) = "D" Then
            tbl_med_comb.Edit
            tbl_med_comb.Delete
        End If
        tbl_med_comb.MoveNext
    Loop
    cmd_sair.SetFocus
End Sub
Private Sub Finaliza()
    tbl_combustivel.Close
    tbl_med_comb.Close
End Sub
Private Sub Inclui()
    lOpcao = 1
    DesativaBotoes
    cmd_novo.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    'If Not (tbl_med_comb.BOF And tbl_med_comb.EOF) Then
    '    tbl_med_comb.Seek "<", g_empresa, "99/99/99", "ZZ"
    '    If Not tbl_med_comb.NoMatch Then
    '        If tbl_med_comb!empresa = g_empresa Then
    '            txt_numero = tbl_med_comb!numero + 1
    '        End If
    '    End If
    'End If
End Sub
Private Sub cbo_combustivel_GotFocus()
    SendMessageLong cbo_combustivel.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_combustivel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_tanque_1.SetFocus
    End If
End Sub
Private Sub cbo_combustivel_LostFocus()
    If cbo_combustivel.ListIndex = -1 Then
        cbo_combustivel.SetFocus
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
    msk_tanque_1.SetFocus
End Sub
Private Sub cmd_anterior_Click()
    Dim data_x As Date
    Dim x_flag As Integer
    If Not (tbl_med_comb.BOF And tbl_med_comb.EOF) Then
        tbl_med_comb.MovePrevious
        If tbl_med_comb.BOF Then
            x_flag = 1
        Else
            If tbl_med_comb!Empresa <> g_empresa Then
                x_flag = 1
            End If
        End If
        If x_flag = 1 Then
            MsgBox "Início de Arquivo.", 48, "Atenção!"
            tbl_med_comb.MoveNext
            cmd_proximo.SetFocus
        Else
            AtualTela
        End If
    End If
End Sub
Private Sub cmd_cancelar_Click()
    LimpaTela
    If BuscaRegistro(lData, lTipoCombustivel) Then
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
    msk_data = "__/__/____"
    cbo_combustivel.ListIndex = -1
    msk_tanque_1 = ""
    msk_tanque_2 = ""
    msk_tanque_3 = ""
    msk_tanque_4 = ""
    msk_tanque_5 = ""
    msk_tanque_6 = ""
    txt_observacao_1 = ""
    txt_observacao_2 = ""
    txt_observacao_3 = ""
End Sub
Private Sub cmd_excluir_Click()
    If tbl_med_comb!Data <> "" Then
        If (MsgBox("Deseja Realmente Excluir Este Registro?", 4 + 32 + 256, "Exclusão de Registro!")) = 6 Then
            tbl_med_comb.Edit
            tbl_med_comb.Delete
            LimpaTela
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
    Inclui
    frm_dados.Enabled = True
    msk_data.SetFocus
End Sub
Private Sub cmd_ok_Click()
    On Error GoTo FileError
    If ValidaCampos Then
        AtivaBotoes
        If lOpcao = 1 Then
            tbl_med_comb.AddNew
            AtualTabe
            tbl_med_comb.Update
            cmd_novo.SetFocus
        ElseIf lOpcao = 2 Then
            tbl_med_comb.Edit
            AtualTabe
            tbl_med_comb.Update
            cmd_novo.SetFocus
        End If
        Call BuscaRegistro(lData, lTipoCombustivel)
    End If
    Exit Sub
FileError:
    ErroArquivo tbl_med_comb.Name, "Medição de Combustívela"
    Exit Sub
End Sub
Private Sub TransfereDadosLMC()
    On Error GoTo FileError
    Dim x_data As Date
    Dim tbl_med_comb_normal As Table
    x_data = CDate("01/01/1900")
    'Busca ultima data com movimento
    If tbl_med_comb.RecordCount > 0 Then
        tbl_med_comb.Seek "<", g_empresa, CDate("31/12/2500"), "ZZ"
        If Not tbl_med_comb.NoMatch Then
            If tbl_med_comb!Empresa = g_empresa Then
                x_data = tbl_med_comb!Data
            End If
        End If
    End If
    x_data = x_data + 1
    If (MsgBox("Na empresa " & g_nome_empresa & Chr(10) & "Será transferido a medição de combustível apartir da data " & x_data & "." & Chr(10) & Chr(10) & "Deseja realmente fazer esta transferência?", vbYesNo + 256, "Transfere a Medição de Combustível Para o L.M.C.!")) = 7 Then
        Exit Sub
    End If
    'Transfere Dados para o LMC
    Set tbl_med_comb_normal = bd_sgp.OpenTable("Medicao_Combustivel")
    With tbl_med_comb_normal
        If .RecordCount > 0 Then
            .Index = "id_data"
            .Seek ">", g_empresa, CDate(x_data), "  "
            If Not .NoMatch Then
                Do Until .EOF
                    If !Empresa <> g_empresa Then
                        Exit Do
                    End If
                    tbl_med_comb.AddNew
                    tbl_med_comb!Empresa = !Empresa
                    tbl_med_comb!Data = !Data
                    tbl_med_comb!tipo_combustivel = !tipo_combustivel
                    tbl_med_comb!tanque_1 = !tanque_1
                    tbl_med_comb!tanque_2 = !tanque_2
                    tbl_med_comb!tanque_3 = !tanque_3
                    tbl_med_comb!tanque_4 = !tanque_4
                    tbl_med_comb!tanque_5 = !tanque_5
                    tbl_med_comb!tanque_6 = !tanque_6
                    tbl_med_comb!observacao_1 = !observacao_1
                    tbl_med_comb!observacao_2 = !observacao_2
                    tbl_med_comb!observacao_3 = !observacao_3
                    tbl_med_comb.Update
                    .MoveNext
                Loop
                MsgBox "A empresa " & g_nome_empresa & Chr(10) & "Está com a medição de combustível transferida para o L.M.C.", vbInformation, "Transferência Concluida!"
            Else
                MsgBox "A empresa " & g_nome_empresa & Chr(10) & "Não tem medição de combustível à ser transferida para o L.M.C.", vbInformation, "Transferência Não Concluida!"
            End If
        End If
    End With
    Exit Sub
FileError:
    ErroArquivo tbl_med_comb.Name, "Medição de Combustívela"
    Resume Next
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If Not IsDate(msk_data) Then
        MsgBox "Informe a Data da Medição.", 64, "Atenção!"
        msk_data.SetFocus
    ElseIf Not lTipoCombustivel > "" Then
        MsgBox "Informe o Combustível.", 64, "Atenção!"
        cbo_combustivel.SetFocus
    ElseIf fValidaValor1(msk_tanque_1) = 0 And fValidaValor1(msk_tanque_2) = 0 And fValidaValor1(msk_tanque_3) = 0 And fValidaValor1(msk_tanque_4) = 0 And fValidaValor1(msk_tanque_5) = 0 And fValidaValor1(msk_tanque_6) = 0 Then
        MsgBox "Informe o Estoque de Algum Tanque.", 64, "Atenção!"
        msk_tanque_1.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_pesquisa_Click()
    consulta_medicao_combustiveis.Show 1
    If Len(g_string) > 0 Then
        lData = RetiraGString(1)
        lTipoCombustivel = RetiraGString(2)
        Call BuscaRegistro(lData, lTipoCombustivel)
    End If
End Sub
Private Sub cmd_primeiro_Click()
    With tbl_med_comb
        .Seek ">", g_empresa, CDate("01/01/1900"), "ZZ"
        If Not .NoMatch Then
            If !Empresa = g_empresa Then
                AtualTela
                cmd_proximo.SetFocus
                Exit Sub
            End If
        End If
        MsgBox "Não há registros nesta empresa.", 64, "Erro de Verificação!"
    End With
End Sub
Private Sub cmd_proximo_Click()
    Dim data_x As Date
    data_x = "31/12/99"
    Dim x_flag As Integer
    If Not (tbl_med_comb.BOF And tbl_med_comb.EOF) Then
        tbl_med_comb.MoveNext
        If tbl_med_comb.EOF Then
            x_flag = 1
        Else
            If tbl_med_comb!Empresa <> g_empresa Then
                x_flag = 1
            End If
        End If
        If x_flag = 1 Then
            MsgBox "Fim de Arquivo.", 48, "Atenção!"
            tbl_med_comb.MovePrevious
            cmd_anterior.SetFocus
        Else
            AtualTela
        End If
    End If
End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub cmd_transfere_dados_lmc_Click()
    TransfereDadosLMC
    cmd_cancelar_Click
End Sub
Private Sub cmd_ultimo_Click()
    With tbl_med_comb
        .Seek "<", g_empresa, CDate("31/12/2500"), "ZZ"
        If Not .NoMatch Then
            If !Empresa = g_empresa Then
                AtualTela
                cmd_anterior.SetFocus
                Exit Sub
            End If
        End If
        MsgBox "Não há registros nesta empresa.", 64, "Erro de Verificação!"
    End With
End Sub
Private Sub Command1_Click()
    'GambiDeletaMedicao
End Sub
Private Sub Form_Activate()
    DesativaBotoes
    If BuscaDados Then
        AtivaBotoes
    Else
        cmd_novo.Enabled = True
        cmd_sair.Enabled = True
    End If
    If cmd_novo.Enabled Then
        cmd_novo.SetFocus
    End If
End Sub
Private Sub Form_Load()
    Screen.MousePointer = 1
    CentraForm Me
    Set tbl_combustivel = bd_sgp.OpenTable("Combustivel")
    If g_nome_usuario = "L.M.C." Then
        Set tbl_med_comb = bd_sgp.OpenTable("Medicao_Combustivel_LMC")
        Me.Caption = Me.Caption & " - LMC"
    Else
        Set tbl_med_comb = bd_sgp.OpenTable("Medicao_Combustivel")
    End If
    PreencheCboCombustivel
    tbl_med_comb.Index = "id_data"
End Sub
Private Sub PreencheCboCombustivel()
    cbo_combustivel.Clear
    With tbl_combustivel
        If .RecordCount > 0 Then
            .Index = "id_codigo"
            .Seek ">=", g_empresa, "  "
            Do Until .EOF
                If !Empresa <> g_empresa Then
                    Exit Do
                End If
                cbo_combustivel.AddItem !Codigo & " - " & !Nome
                .MoveNext
            Loop
        End If
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_GotFocus()
    If Not IsDate(msk_data) Then
        msk_data = g_data_def
    End If
End Sub
Private Sub msk_data_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_combustivel.SetFocus
    End If
End Sub
Private Sub msk_tanque_1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        msk_tanque_2.SetFocus
    End If
End Sub
Private Sub msk_tanque_1_LostFocus()
    msk_tanque_1 = Format(msk_tanque_1, "###,##0.0")
End Sub
Private Sub msk_tanque_2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        msk_tanque_3.SetFocus
    End If
End Sub
Private Sub msk_tanque_2_LostFocus()
    msk_tanque_2 = Format(msk_tanque_2, "###,##0.0")
End Sub
Private Sub msk_tanque_3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        msk_tanque_4.SetFocus
    End If
End Sub
Private Sub msk_tanque_3_LostFocus()
    msk_tanque_3 = Format(msk_tanque_3, "###,##0.0")
End Sub
Private Sub msk_tanque_4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        msk_tanque_5.SetFocus
    End If
End Sub
Private Sub msk_tanque_4_LostFocus()
    msk_tanque_4 = Format(msk_tanque_4, "###,##0.0")
End Sub
Private Sub msk_tanque_5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        msk_tanque_6.SetFocus
    End If
End Sub
Private Sub msk_tanque_5_LostFocus()
    msk_tanque_5 = Format(msk_tanque_5, "###,##0.0")
End Sub
Private Sub msk_tanque_6_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_observacao_1.SetFocus
    End If
End Sub
Private Sub msk_tanque_6_LostFocus()
    msk_tanque_6 = Format(msk_tanque_6, "###,##0.0")
End Sub
Private Sub txt_observacao_1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_observacao_2.SetFocus
    End If
End Sub
Private Sub txt_observacao_2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_observacao_3.SetFocus
    End If
End Sub
Private Sub txt_observacao_3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
