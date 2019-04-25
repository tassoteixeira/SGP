VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form baixa_cartao 
   Caption         =   "Baixa de Cartão de Crédito"
   ClientHeight    =   3675
   ClientLeft      =   1920
   ClientTop       =   2790
   ClientWidth     =   5055
   Icon            =   "baixa_cartao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "baixa_cartao.frx":030A
   ScaleHeight     =   3675
   ScaleWidth      =   5055
   Begin VB.Frame frm_dados 
      Height          =   2595
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   4815
      Begin VB.TextBox txtTaxaAdministrativa 
         Height          =   285
         Left            =   2880
         MaxLength       =   6
         TabIndex        =   15
         Top             =   2160
         Width           =   615
      End
      Begin VB.CommandButton cmdDataAntecipacao 
         Height          =   315
         Left            =   1260
         Picture         =   "baixa_cartao.frx":0750
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   2160
         Width           =   495
      End
      Begin VB.OptionButton opt_vencimento 
         Caption         =   "Data de &vencimento"
         Height          =   195
         Left            =   2880
         TabIndex        =   2
         Top             =   300
         Width           =   1755
      End
      Begin VB.OptionButton opt_emissao 
         Caption         =   "Data de e&missão"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.ComboBox cbo_cartao_credito 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1500
         Width           =   3435
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   1260
         Picture         =   "baixa_cartao.frx":1A2A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   840
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   4020
         Picture         =   "baixa_cartao.frx":2D04
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   840
         Width           =   495
      End
      Begin MSMask.MaskEdBox msk_data_final 
         Height          =   300
         Left            =   2880
         TabIndex        =   7
         Top             =   840
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_data_inicial 
         Height          =   300
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskDataAntecipacao 
         Height          =   300
         Left            =   120
         TabIndex        =   12
         Top             =   2160
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         Caption         =   "&Taxa de Antecipação"
         Height          =   255
         Left            =   2880
         TabIndex        =   14
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Data da a&ntecipação"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   1920
         Width           =   1635
      End
      Begin VB.Label Label1 
         Caption         =   "&Cartão"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   1260
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "&Data inicial"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "D&ata final"
         Height          =   255
         Index           =   2
         Left            =   2880
         TabIndex        =   6
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmd_baixa 
      Caption         =   "&Baixar"
      Height          =   855
      Left            =   240
      Picture         =   "baixa_cartao.frx":3FDE
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Baixa os cartões no período informado."
      Top             =   2760
      Width           =   795
   End
   Begin VB.CommandButton cmd_estornar 
      Caption         =   "&Estornar"
      Height          =   855
      Left            =   2160
      Picture         =   "baixa_cartao.frx":52B8
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Estorna os cartões no período informado."
      Top             =   2760
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4020
      Picture         =   "baixa_cartao.frx":6592
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   2760
      Width           =   795
   End
End
Attribute VB_Name = "baixa_cartao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lTotal As Currency
Dim lSQL As String

Dim rst_cartao As New adodb.Recordset
Private IntegracaoCaixa As New cIntegracaoCaixa
Private MovCartaoCredito As New cMovimentoCartaoCredito
Private MovCaixa As New cMovimentoCaixa
Private Sub AtivaBotoes(xAtiva As Boolean)
    cmd_baixa.Enabled = xAtiva
    cmd_estornar.Enabled = xAtiva
    cmd_sair.Enabled = xAtiva
End Sub
Private Sub PreencheCboCartaoCredito()
    cbo_cartao_credito.Clear
    cbo_cartao_credito.AddItem "Todos os Cartões"
    cbo_cartao_credito.ItemData(cbo_cartao_credito.NewIndex) = 0
    'Set rst_cartao = cnnSGP.Execute("Select Codigo, Nome From Cartao_Credito Order By Nome")
    Set rst_cartao = Conectar.RsConexao("Select Codigo, Nome From Cartao_Credito Order By Nome")
    'Set FldCodigo = g_rs.Fields(0)
    'Set FldNome = g_rs.Fields(1)
    'loop RecordSet
    With rst_cartao
        If Not .BOF Or Not .EOF Then
            .MoveFirst
            Do Until .EOF
                cbo_cartao_credito.AddItem !Nome
                cbo_cartao_credito.ItemData(cbo_cartao_credito.NewIndex) = !Codigo
                .MoveNext
            Loop
        End If
        .Close
    End With
End Sub
Private Sub PreparaBaixa()
    Dim xDataEmissao As Boolean
    Dim xString As String
    
    If opt_emissao Then
        xDataEmissao = True
    Else
        xDataEmissao = False
    End If
    lTotal = MovCartaoCredito.TotalEntreDatas(g_empresa, xDataEmissao, True, CDate(msk_data_inicial.Text), CDate(msk_data_final.Text), Val(cbo_cartao_credito.ItemData(cbo_cartao_credito.ListIndex)))
    If lTotal > 0 Then
        If (MsgBox("No período informado tem R$ " & Format(lTotal, "###,###,##0.00") & " em cartão à ser baixado." & Chr(13) & Chr(10) & Chr(10) & "Deseja realmente baixá-los?", vbQuestion + vbYesNo + vbDefaultButton2, "Baixa de Cartão de Crédito.")) = vbYes Then
            If opt_emissao.Value = True Then
                xString = "Data Emissão:"
            Else
                xString = "Data Vencimento:"
            End If
            xString = xString & msk_data_inicial.Text & " a " & msk_data_final.Text
            xString = xString & " Cartão:" & cbo_cartao_credito.Text
            Call GravaAuditoria(1, Me.name, 18, xString)
            If IsDate(mskDataAntecipacao.Text) Then
                xString = "Data Antecipação:" & mskDataAntecipacao.Text
                xString = xString & " Taxa ADM:" & txtTaxaAdministrativa.Text
                Call GravaAuditoria(2, Me.name, 18, xString)
            End If
            Baixa
        End If
    Else
        MsgBox "Não existe cartão à ser baixado no período informado!", vbExclamation, "Baixa de Cartão de Crédito."
    End If
    AtivaBotoes (True)
    cmd_sair.SetFocus
End Sub
Private Sub PreparaEstorno()
    Dim xDataEmissao As Boolean
    Dim xString As String
    
    If opt_emissao Then
        xDataEmissao = True
    Else
        xDataEmissao = False
    End If
    lTotal = MovCartaoCredito.TotalEntreDatas(g_empresa, xDataEmissao, False, CDate(msk_data_inicial.Text), CDate(msk_data_final.Text), Val(cbo_cartao_credito.ItemData(cbo_cartao_credito.ListIndex)))
    If lTotal > 0 Then
        If (MsgBox("No período informado tem R$ " & Format(lTotal, "###,###,##0.00") & " em cartão baixado à ser estornado." & Chr(13) & Chr(10) & Chr(10) & "Deseja realmente estornar?", vbQuestion + vbYesNo + vbDefaultButton2, "Estorno de Cartão de Crédito.")) = vbYes Then
            If Me.opt_emissao.Value = True Then
                xString = "Data Emissão:"
            Else
                xString = "Data Vencimento:"
            End If
            xString = xString & msk_data_inicial.Text & " a " & msk_data_final.Text
            xString = xString & " Cartão:" & cbo_cartao_credito.Text
            Call GravaAuditoria(1, Me.name, 19, xString)
            Estorno
        End If
    Else
        MsgBox "Não existe cartão baixado à ser estornado no período informado!", vbExclamation, "Estorno de Cartão de Crédito."
    End If
    AtivaBotoes (True)
    cmd_sair.SetFocus
End Sub
Private Sub Baixa()
    Dim xDataEmissao As Boolean
    Dim xDataAntecipacao As Date
    Dim xTaxaAdministracao As Currency
    Dim xMensagem As String
    
    On Error GoTo FileError
    
    If opt_emissao Then
        xDataEmissao = True
    Else
        xDataEmissao = False
    End If
    If IsDate(mskDataAntecipacao.Text) Then
        xDataAntecipacao = CDate(mskDataAntecipacao.Text)
        xTaxaAdministracao = fValidaValor(txtTaxaAdministrativa.Text)
    Else
        xDataAntecipacao = "00:00:00"
        xTaxaAdministracao = 0
    End If
    If MovCartaoCredito.BaixaCartao(g_empresa, g_usuario, xDataEmissao, CDate(msk_data_inicial.Text), CDate(msk_data_final.Text), Val(cbo_cartao_credito.ItemData(cbo_cartao_credito.ListIndex)), xDataAntecipacao, xTaxaAdministracao) Then
        xMensagem = "Baixa de cartão concluida!"
        MsgBox xMensagem, vbExclamation, "Baixa de Cartão de Crédito."
        Call GravaAuditoria(1, Me.name, 22, xMensagem & " Total:" & Format(lTotal, "###,###,##0.00"))
    Else
        xMensagem = "Baixa de cartão não foi concluida!"
        MsgBox xMensagem, vbExclamation, "Erro de Integridade!"
        Call GravaAuditoria(1, Me.name, 22, xMensagem & " Total:" & Format(lTotal, "###,###,##0.00"))
    End If
    Exit Sub
FileError:
    MsgBox "Erro na Baixa: " & Error
    Exit Sub
End Sub
Private Sub Estorno()
    Dim xDataEmissao As Boolean
    Dim xMensagem As String
    
    On Error GoTo FileError
    
    If opt_emissao Then
        xDataEmissao = True
    Else
        xDataEmissao = False
    End If
    If MovCartaoCredito.EstornaCartao(g_empresa, xDataEmissao, CDate(msk_data_inicial.Text), CDate(msk_data_final.Text), Val(cbo_cartao_credito.ItemData(cbo_cartao_credito.ListIndex))) Then
        xMensagem = "Estorno de cartão concluido!"
        MsgBox xMensagem, vbExclamation, "Estorno de Cartão de Crédito."
        Call GravaAuditoria(1, Me.name, 22, xMensagem & " Total:" & Format(lTotal, "###,###,##0.00"))
    Else
        xMensagem = "Estorno de cartão não foi concluido!"
        MsgBox xMensagem, vbExclamation, "Erro de Integridade!"
        Call GravaAuditoria(1, Me.name, 22, xMensagem & " Total:" & Format(lTotal, "###,###,##0.00"))
    End If
    Exit Sub
FileError:
    MsgBox "Erro no Estorno: " & Error
    Exit Sub
End Sub
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set MovCartaoCredito = Nothing
End Sub
Private Sub cbo_cartao_credito_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        mskDataAntecipacao.SetFocus
    End If
End Sub
Private Sub cmd_baixa_Click()
    If ValidaCampos Then
   '     AtivaBotoes (False)
        PreparaBaixa
    End If
End Sub
Private Sub cmd_data_f_Click()
    g_string = msk_data_final.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_inicial.Text = RetiraGString(1)
        msk_data_final.Text = RetiraGString(2)
    Else
        msk_data_final.Text = RetiraGString(1)
    End If
    g_string = ""
    cbo_cartao_credito.SetFocus
End Sub
Private Sub cmd_data_i_Click()
    g_string = msk_data_inicial.Enabled
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_inicial.Text = RetiraGString(1)
        msk_data_final.Text = RetiraGString(2)
        cbo_cartao_credito.SetFocus
    Else
        msk_data_inicial.Text = RetiraGString(1)
        msk_data_final.SetFocus
    End If
    g_string = ""
End Sub
Private Sub cmd_estornar_Click()
    If ValidaCampos Then
        AtivaBotoes (False)
        PreparaEstorno
    End If
End Sub
Function ValidaCampos() As Boolean
    ValidaCampos = False
    If Not IsDate(msk_data_inicial.Text) Then
        MsgBox "Informe a data da inicial.", vbInformation, "Atenção!"
        msk_data_inicial.SetFocus
    ElseIf Not IsDate(msk_data_final.Text) Then
        MsgBox "Informe a data da final.", vbInformation, "Atenção!"
        msk_data_final.SetFocus
    ElseIf CDate(msk_data_final.Text) < CDate(msk_data_inicial.Text) Then
        MsgBox "A data final deve ser maior ou igual a " & msk_data_inicial.Text & ".", vbInformation, "Atenção!"
        msk_data_final.SetFocus
    ElseIf mskDataAntecipacao.Text <> "__/__/____" And Not IsDate(mskDataAntecipacao.Text) Then
        MsgBox "Informe uma data de antecipação válida.", vbInformation, "Atenção!"
        mskDataAntecipacao.SetFocus
    ElseIf IsDate(mskDataAntecipacao.Text) And Val(txtTaxaAdministrativa.Text) = 0 Then
        MsgBox "Informe o percentual da taxa administrativa.", vbInformation, "Atenção!"
        txtTaxaAdministrativa.SetFocus
    ElseIf cbo_cartao_credito.ListIndex = -1 Then
        MsgBox "Selecione um tipo de cartão de crédito.", vbInformation, "Atenção!"
        cbo_cartao_credito.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub BuscaDatas()
    Dim x_campo As String
    msk_data_inicial.Text = CDate(g_data_def - 1)
    msk_data_final.Text = CDate(g_data_def - 1)
    If opt_emissao Then
        x_campo = "[Data de Emissao]"
    Else
        x_campo = "[Data do Vencimento]"
    End If
    'Busca Data Inicial
    lSQL = "SELECT TOP 1 " & x_campo & " AS DataX"
    lSQL = lSQL & " FROM Movimento_Cartao_Credito"
    lSQL = lSQL & " ORDER BY " & x_campo
    Set rst_cartao = Conectar.RsConexao(lSQL)
    If Not rst_cartao.EOF Then
        msk_data_inicial.Text = rst_cartao!DataX
    End If
    'Busca Data Final
    lSQL = "SELECT TOP 1 " & x_campo & " AS DataX"
    lSQL = lSQL & " FROM Movimento_Cartao_Credito"
    lSQL = lSQL & " ORDER BY " & x_campo & " DESC"
    Set rst_cartao = Conectar.RsConexao(lSQL)
    If Not rst_cartao.EOF Then
        msk_data_final.Text = rst_cartao!DataX
    End If
    mskDataAntecipacao.Text = "__/__/____"
    txtTaxaAdministrativa.Text = ""
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    BuscaDatas
    cbo_cartao_credito.SetFocus
    Screen.MousePointer = 1
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        cmd_sair_Click
    End If
End Sub
Private Sub Form_Load()
    CentraForm Me
    PreencheCboCartaoCredito
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_final_GotFocus()
    msk_data_final.SelStart = 0
    msk_data_final.SelLength = 5
End Sub
Private Sub msk_data_final_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_cartao_credito.SetFocus
    End If
End Sub
Private Sub msk_data_inicial_GotFocus()
    msk_data_inicial.SelStart = 0
    msk_data_inicial.SelLength = 5
End Sub
Private Sub msk_data_inicial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data_final.SetFocus
    End If
End Sub
Private Sub mskDataAntecipacao_GotFocus()
    mskDataAntecipacao.SelStart = 0
    mskDataAntecipacao.SelLength = 5
End Sub
Private Sub mskDataAntecipacao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtTaxaAdministrativa.SetFocus
    End If
End Sub
Private Sub opt_emissao_Click()
    BuscaDatas
End Sub
Private Sub opt_vencimento_Click()
    BuscaDatas
End Sub
Private Sub txtTaxaAdministrativa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_baixa.SetFocus
    End If
End Sub
Private Sub txtTaxaAdministrativa_LostFocus()
    txtTaxaAdministrativa.Text = Format(txtTaxaAdministrativa.Text, "##0.00")
End Sub
