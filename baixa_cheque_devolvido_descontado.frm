VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form baixa_cheque_devolvido_descontado 
   Caption         =   "Baixa de Cheque Devolvido Descontado"
   ClientHeight    =   5175
   ClientLeft      =   2325
   ClientTop       =   3105
   ClientWidth     =   6975
   Icon            =   "baixa_cheque_devolvido_descontado.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "baixa_cheque_devolvido_descontado.frx":030A
   ScaleHeight     =   5175
   ScaleWidth      =   6975
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   180
      Picture         =   "baixa_cheque_devolvido_descontado.frx":0750
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Cria um novo registro."
      Top             =   4200
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   1080
      Picture         =   "baixa_cheque_devolvido_descontado.frx":1DE2
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Altera o registro atual."
      Top             =   4200
      Width           =   795
   End
   Begin VB.CommandButton cmd_pesquisa 
      Caption         =   "&Pesquisa"
      Height          =   855
      Left            =   2880
      Picture         =   "baixa_cheque_devolvido_descontado.frx":32DC
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Pesquisa um registro específico."
      Top             =   4200
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3780
      Picture         =   "baixa_cheque_devolvido_descontado.frx":474E
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   4200
      Width           =   795
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Extornar"
      Height          =   855
      Left            =   1980
      Picture         =   "baixa_cheque_devolvido_descontado.frx":5DE0
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Extorna o registro atual."
      Top             =   4200
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Enabled         =   0   'False
      Height          =   4035
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   6735
      Begin VB.TextBox txt_funcionario 
         Height          =   285
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   16
         Top             =   2940
         Width           =   4935
      End
      Begin VB.CheckBox chk_inativo 
         Height          =   285
         Left            =   1680
         TabIndex        =   20
         Top             =   3660
         Width           =   285
      End
      Begin MSMask.MaskEdBox msk_data_pagamento 
         Height          =   315
         Left            =   1680
         TabIndex        =   18
         Top             =   3300
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lbl_motivo 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1680
         TabIndex        =   14
         Top             =   2400
         Width           =   2535
      End
      Begin VB.Label lbl_data_entrega 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1680
         TabIndex        =   12
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label lbl_emitente 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1680
         TabIndex        =   10
         Top             =   1680
         Width           =   4935
      End
      Begin VB.Label lbl_valor 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1680
         TabIndex        =   8
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lbl_cheque 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1680
         TabIndex        =   6
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lbl_agencia 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1680
         TabIndex        =   4
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lbl_banco 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1680
         TabIndex        =   2
         Top             =   240
         Width           =   375
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   6720
         Y1              =   2820
         Y2              =   2820
      End
      Begin VB.Label Label6 
         Caption         =   "Data do Pagamento"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Inativo"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   3660
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Motivo"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Funcionário"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Número do Banco"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Nome do Emitente"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Data da Entrega"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Número do Cheque"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Número da Agência"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Valor do Cheque"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   1455
      End
   End
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   4680
      TabIndex        =   28
      Top             =   4080
      Width           =   2175
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "baixa_cheque_devolvido_descontado.frx":70BA
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "baixa_cheque_devolvido_descontado.frx":85B4
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "baixa_cheque_devolvido_descontado.frx":9AAE
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "baixa_cheque_devolvido_descontado.frx":AF20
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
      Left            =   6060
      Picture         =   "baixa_cheque_devolvido_descontado.frx":C4A2
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Cancela o registro atual."
      Top             =   4200
      Width           =   795
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   5160
      Picture         =   "baixa_cheque_devolvido_descontado.frx":D99C
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Confirma o registro atual."
      Top             =   4200
      Width           =   795
   End
End
Attribute VB_Name = "baixa_cheque_devolvido_descontado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flag_baixa_cheque_devolvido_descontado As Integer
Dim lOpcao As Integer
Dim l_empresa As Integer
Dim l_data_pagamento As Date
Dim l_banco As Integer
Dim l_agencia As String
Dim l_cheque As String
Dim l_codigo_funcionario As Integer

Private Funcionario As New cFuncionario

Dim tbl_baixa_cheque_devolvido_descontado As Table
Dim tbl_baixa_cheque_devolvido As Table
Private Sub AtualTabe()
    With tbl_baixa_cheque_devolvido_descontado
        !Empresa = g_empresa
        ![Numero do Banco] = Format(lbl_banco, "000")
        ![Numero da Agencia] = Format(lbl_agencia, "0000")
        ![Numero do Cheque] = Format(lbl_cheque, "000000")
        !valor = fValidaValor2(lbl_valor)
        !Emitente = lbl_emitente
        ![Data da Entrega] = CDate(lbl_data_entrega)
        !Motivo = lbl_motivo
        ![Codigo do Funcionario] = l_codigo_funcionario
        ![Nome do Funcionario] = Mid(txt_funcionario, 1, 30)
        ![Data do Pagamento] = msk_data_pagamento
        If chk_inativo Then
            !Inativo = True
        Else
            !Inativo = False
        End If
        l_data_pagamento = ![Data do Pagamento]
        l_banco = ![Numero do Banco]
        l_agencia = ![Numero da Agencia]
        l_cheque = ![Numero do Cheque]
        l_codigo_funcionario = ![Codigo do Funcionario]
    End With
End Sub
Private Sub AtualTabeBaixaChequeDevolvido()
    'BaixaChequeDevolvido.Empresa = g_empresa
    With tbl_baixa_cheque_devolvido
        !Empresa = g_empresa
        ![Numero do Banco] = Format(lbl_banco, "000")
        ![Numero da Agencia] = Format(lbl_agencia, "0000")
        ![Numero do Cheque] = Format(lbl_cheque, "000000")
        !valor = fValidaValor2(lbl_valor)
        !Emitente = lbl_emitente
        ![Data da Entrega] = CDate(lbl_data_entrega)
        !Motivo = lbl_motivo
        ![Codigo do Funcionario] = l_codigo_funcionario
        ![Data do Pagamento] = tbl_baixa_cheque_devolvido_descontado![Data da Devolucao]
        ![Valor Pago] = tbl_baixa_cheque_devolvido_descontado!valor
        If chk_inativo Then
            !Inativo = True
        Else
            !Inativo = False
        End If
    End With
End Sub
Function BuscaDados() As Boolean
    BuscaDados = False
    With tbl_baixa_cheque_devolvido_descontado
        If .RecordCount > 0 Then
            If lOpcao = 3 Then
                If Not .EOF Then
                    .MoveNext
                    If Not .EOF Then
                        If !Empresa = g_empresa Then
                            AtualTela
                            BuscaDados = True
                            Exit Function
                        End If
                    End If
                End If
            End If
            .Seek "<", g_empresa, CDate("31/12/2500"), 9999, "ZZZZ", "ZZZZZZ"
            If Not .NoMatch Then
                If !Empresa = g_empresa Then
                    AtualTela
                    BuscaDados = True
                    Exit Function
                End If
            End If
        End If
        LimpaTela
    End With
End Function
Function BuscaRegistro(x_data As Date, x_banco As Integer, x_agencia As String, x_cheque As String) As Boolean
    BuscaRegistro = False
    If tbl_baixa_cheque_devolvido_descontado.RecordCount > 0 Then
        tbl_baixa_cheque_devolvido_descontado.Seek "=", g_empresa, x_data, x_banco, x_agencia, x_cheque
        If Not tbl_baixa_cheque_devolvido_descontado.NoMatch Then
            AtualTela
            BuscaRegistro = True
        End If
    End If
End Function
Function BuscaRegistroBaixaChequeDevolvido(x_data As Date, x_banco As Integer, x_agencia As String, x_cheque As String) As Boolean
    BuscaRegistroBaixaChequeDevolvido = False
    With tbl_baixa_cheque_devolvido
        If .RecordCount > 0 Then
            .Seek "=", g_empresa, x_data, x_banco, x_agencia, x_cheque
            If Not .NoMatch Then
                l_codigo_funcionario = ![Codigo do Funcionario]
                lbl_banco = Format(![Numero do Banco], "000")
                lbl_agencia = Format(![Numero da Agencia], "0000")
                lbl_cheque = Format(![Numero do Cheque], "000000")
                lbl_valor = Format(!valor, "###,##0.00")
                lbl_emitente = !Emitente
                lbl_data_entrega = Format(![Data da Entrega], "dd/mm/yyyy")
                lbl_motivo = !Motivo
                If Funcionario.LocalizarCodigo(g_empresa, ![Codigo do Funcionario]) Then
                    txt_funcionario = Funcionario.Nome
                End If
                If !Inativo Then
                    chk_inativo.Value = 1
                Else
                    chk_inativo.Value = 0
                End If
                BuscaRegistroBaixaChequeDevolvido = True
            End If
        End If
    End With
End Function
Private Sub AtualTela()
    With tbl_baixa_cheque_devolvido_descontado
        l_data_pagamento = ![Data do Pagamento]
        l_banco = ![Numero do Banco]
        l_agencia = ![Numero da Agencia]
        l_cheque = ![Numero do Cheque]
        l_codigo_funcionario = ![Codigo do Funcionario]
        lbl_banco = Format(![Numero do Banco], "000")
        lbl_agencia = Format(![Numero da Agencia], "0000")
        lbl_cheque = Format(![Numero do Cheque], "000000")
        lbl_valor = Format(!valor, "###,##0.00")
        lbl_emitente = !Emitente
        lbl_data_entrega = Format(![Data da Entrega], "dd/mm/yyyy")
        lbl_motivo = !Motivo
        txt_funcionario = ![Nome do Funcionario]
        If !Inativo Then
            chk_inativo.Value = 1
        Else
            chk_inativo.Value = 0
        End If
        msk_data_pagamento = Format(![Data do Pagamento], "dd/mm/yyyy")
    End With
    frm_dados.Enabled = False
End Sub
Private Sub Finaliza()
    Set Funcionario = Nothing
    tbl_baixa_cheque_devolvido_descontado.Close
    tbl_baixa_cheque_devolvido.Close
End Sub
Private Sub Inclui()
    lOpcao = 1
    DesativaBotoes
    cmd_novo.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
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
    msk_data_pagamento.SetFocus
End Sub
Private Sub cmd_anterior_Click()
    If tbl_baixa_cheque_devolvido_descontado.RecordCount > 0 Then
        tbl_baixa_cheque_devolvido_descontado.MovePrevious
        If Not tbl_baixa_cheque_devolvido_descontado.BOF Then
            If tbl_baixa_cheque_devolvido_descontado!Empresa = g_empresa Then
                AtualTela
                Exit Sub
            End If
        End If
        MsgBox "Início de Arquivo.", 48, "Atenção!"
        tbl_baixa_cheque_devolvido_descontado.MoveNext
        cmd_proximo.SetFocus
    End If
End Sub
Private Sub cmd_cancelar_Click()
    LimpaTela
    If BuscaRegistro(l_data_pagamento, l_banco, l_agencia, l_cheque) Then
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
    lbl_banco = ""
    lbl_agencia = ""
    lbl_cheque = ""
    lbl_valor = ""
    lbl_emitente = ""
    lbl_data_entrega = "__/__/____"
    lbl_motivo = ""
    txt_funcionario = ""
    msk_data_pagamento = "__/__/____"
    chk_inativo.Value = False
End Sub
Private Sub cmd_excluir_Click()
    If tbl_baixa_cheque_devolvido_descontado![Numero do Cheque] > 0 Then
        If (MsgBox("Deseja Realmente Extornar Este Registro?", 4 + 32 + 256, "Extorno de Registro!")) = 6 Then
            tbl_baixa_cheque_devolvido.AddNew
            AtualTabeBaixaChequeDevolvido
            tbl_baixa_cheque_devolvido.Update
            tbl_baixa_cheque_devolvido_descontado.Edit
            tbl_baixa_cheque_devolvido_descontado.Delete
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
    frm_dados.Enabled = True
    Inclui
    LimpaTela
    consulta_cheque_devolvido_baixado.Show 1
    If Len(g_string) > 0 Then
        If BuscaRegistroBaixaChequeDevolvido(RetiraGString(1), RetiraGString(2), RetiraGString(3), RetiraGString(4)) Then
            txt_funcionario.SetFocus
            Exit Sub
        End If
    End If
    cmd_cancelar_Click
End Sub
Private Sub cmd_ok_Click()
    On Error GoTo FileError
    If ValidaCampos Then
        AtivaBotoes
        If lOpcao = 1 Then
            tbl_baixa_cheque_devolvido_descontado.AddNew
            AtualTabe
            tbl_baixa_cheque_devolvido_descontado.[Data da Devolucao] = tbl_baixa_cheque_devolvido![Data do Pagamento]
            tbl_baixa_cheque_devolvido_descontado.Update
            tbl_baixa_cheque_devolvido.Edit
            tbl_baixa_cheque_devolvido.Delete
        ElseIf lOpcao = 2 Then
            tbl_baixa_cheque_devolvido_descontado.Edit
            AtualTabe
            tbl_baixa_cheque_devolvido_descontado.Update
        End If
        Call BuscaRegistro(l_data_pagamento, l_banco, l_agencia, l_cheque)
        lOpcao = 0
        cmd_novo.SetFocus
    End If
    Exit Sub
FileError:
    ErroArquivo tbl_baixa_cheque_devolvido_descontado.name, "Cheque Devolvido Baixado Descontadoo"
    Exit Sub
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If txt_funcionario = "" Then
        MsgBox "Informe o nome do funcionário.", 64, "Atenção!"
        txt_funcionario.SetFocus
    ElseIf Not IsDate(msk_data_pagamento) Then
        MsgBox "Informe a data do pagamento.", 64, "Atenção!"
        msk_data_pagamento.SetFocus
    ElseIf CDate(msk_data_pagamento) < CDate(lbl_data_entrega) Then
        MsgBox "A data do pagamento deve ser maior ou igual que " & lbl_data_entrega & ".", 64, "Atenção!"
        msk_data_pagamento.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_pesquisa_Click()
    consulta_baixa_cheque_devolvido_descontado.Show 1
    If Len(g_string) > 0 Then
        l_data_pagamento = RetiraGString(1)
        l_banco = RetiraGString(2)
        l_agencia = RetiraGString(3)
        l_cheque = RetiraGString(4)
        Call BuscaRegistro(l_data_pagamento, l_banco, l_agencia, l_cheque)
    End If
End Sub
Private Sub cmd_primeiro_Click()
    If tbl_baixa_cheque_devolvido_descontado.RecordCount > 0 Then
        tbl_baixa_cheque_devolvido_descontado.Seek ">", g_empresa, CDate("01/01/1900"), 0, "    ", "      "
        If Not tbl_baixa_cheque_devolvido_descontado.NoMatch Then
            If tbl_baixa_cheque_devolvido_descontado!Empresa = g_empresa Then
                AtualTela
                cmd_proximo.SetFocus
                Exit Sub
            End If
        End If
        MsgBox "Não há registros nesta empresa.", 64, "Erro de Verificação!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    If tbl_baixa_cheque_devolvido_descontado.RecordCount > 0 Then
        tbl_baixa_cheque_devolvido_descontado.MoveNext
        If Not tbl_baixa_cheque_devolvido_descontado.EOF Then
            If tbl_baixa_cheque_devolvido_descontado!Empresa = g_empresa Then
                AtualTela
                Exit Sub
            End If
        End If
        MsgBox "Fim de Arquivo.", 48, "Atenção!"
        tbl_baixa_cheque_devolvido_descontado.MovePrevious
        cmd_anterior.SetFocus
    End If
End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub cmd_ultimo_Click()
    If tbl_baixa_cheque_devolvido_descontado.RecordCount > 0 Then
        tbl_baixa_cheque_devolvido_descontado.Seek "<", g_empresa, CDate("31/12/2500"), 9999, "ZZZZ", "ZZZZZZ"
        If Not tbl_baixa_cheque_devolvido_descontado.NoMatch Then
            If tbl_baixa_cheque_devolvido_descontado!Empresa = g_empresa Then
                AtualTela
                cmd_anterior.SetFocus
                Exit Sub
            End If
        End If
        MsgBox "Não há registros nesta empresa.", 64, "Erro de Verificação!"
    End If
End Sub
Private Sub Form_Activate()
    If g_empresa <> l_empresa Then
        flag_baixa_cheque_devolvido_descontado = 0
    End If
    If flag_baixa_cheque_devolvido_descontado = 0 Then
        lOpcao = 0
        l_empresa = g_empresa
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
    Else
        flag_baixa_cheque_devolvido_descontado = 0
    End If
    Screen.MousePointer = 1
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
    flag_baixa_cheque_devolvido_descontado = 1
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 And lOpcao = 0 Then
        KeyCode = 0
        cmd_novo_Click
    ElseIf KeyCode = vbKeyF3 And lOpcao = 0 Then
        KeyCode = 0
        cmd_alterar_Click
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
    
    Set tbl_baixa_cheque_devolvido_descontado = bd_sgp.OpenTable("Baixa_Cheque_Devolvido_Descontado")
    Set tbl_baixa_cheque_devolvido = bd_sgp.OpenTable("Baixa_Cheque_Devolvido")
    tbl_baixa_cheque_devolvido_descontado.Index = "id_data_pagamento"
    tbl_baixa_cheque_devolvido.Index = "id_data_pagamento"
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_pagamento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        chk_inativo.SetFocus
    End If
End Sub
Private Sub txt_funcionario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data_pagamento.SetFocus
    End If
End Sub
