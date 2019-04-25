VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form movimento_bancario 
   Appearance      =   0  'Flat
   Caption         =   "Movimentação Bancária"
   ClientHeight    =   4335
   ClientLeft      =   1650
   ClientTop       =   3690
   ClientWidth     =   9030
   ForeColor       =   &H80000008&
   Icon            =   "mov_bancario.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4335
   ScaleWidth      =   9030
   Begin VB.Frame Frame1 
      Height          =   675
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   8895
      Begin MSAdodcLib.Adodc adodc_conta 
         Height          =   330
         Left            =   3570
         Top             =   240
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
         Caption         =   "adodc_conta"
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
      Begin MSDataListLib.DataCombo dtcbo_conta 
         Bindings        =   "mov_bancario.frx":030A
         Height          =   315
         Left            =   1590
         TabIndex        =   2
         Top             =   240
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   "dtcbo_conta"
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         Caption         =   "Co&nta"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1425
      End
   End
   Begin VB.Frame frm_dados 
      Height          =   2715
      Left            =   60
      TabIndex        =   3
      Top             =   600
      Width           =   8895
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2760
         Picture         =   "mov_bancario.frx":0324
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txt_documento 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1620
         MaxLength       =   8
         TabIndex        =   16
         Top             =   2280
         Width           =   1515
      End
      Begin VB.TextBox txt_complemento 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1620
         MaxLength       =   20
         TabIndex        =   14
         Top             =   1890
         Width           =   4035
      End
      Begin VB.TextBox txt_valor 
         Height          =   315
         Left            =   1620
         MaxLength       =   9
         TabIndex        =   8
         Top             =   660
         Width           =   1395
      End
      Begin VB.ComboBox cbo_operacao 
         Height          =   315
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1500
         Width           =   495
      End
      Begin MSMask.MaskEdBox msk_data 
         Height          =   315
         Left            =   1620
         TabIndex        =   5
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
      Begin MSAdodcLib.Adodc adodc_historico 
         Height          =   330
         Left            =   3600
         Top             =   1020
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
         Caption         =   "adodc_historico"
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
      Begin MSDataListLib.DataCombo dtcbo_historico 
         Bindings        =   "mov_bancario.frx":15FE
         Height          =   315
         Left            =   1620
         TabIndex        =   10
         Top             =   1080
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   "dtcbo_historico"
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         Caption         =   "&Data do Movimento"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1425
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         Caption         =   "&Nr. Documento"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   150
         TabIndex        =   15
         Top             =   2280
         Width           =   1425
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         Caption         =   "Co&mplemento"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   150
         TabIndex        =   13
         Top             =   1920
         Width           =   1425
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         Caption         =   "&Histórico"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   150
         TabIndex        =   9
         Top             =   1080
         Width           =   1425
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         Caption         =   "&Valor"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   150
         TabIndex        =   7
         Top             =   660
         Width           =   1425
      End
      Begin VB.Label Label7 
         Caption         =   "O&peração"
         Height          =   315
         Left            =   150
         TabIndex        =   11
         Top             =   1500
         Width           =   1395
      End
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3660
      Picture         =   "mov_bancario.frx":161C
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   3420
      Width           =   795
   End
   Begin VB.CommandButton cmd_pesquisa 
      Caption         =   "&Pesquisa"
      Height          =   855
      Left            =   2760
      Picture         =   "mov_bancario.frx":2CAE
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Pesquisa um registro específico."
      Top             =   3420
      Width           =   795
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   1860
      Picture         =   "mov_bancario.frx":4120
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Exclui o registro atual."
      Top             =   3420
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   960
      Picture         =   "mov_bancario.frx":57B2
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Altera o registro atual."
      Top             =   3420
      Width           =   795
   End
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   60
      Picture         =   "mov_bancario.frx":6CAC
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Cria um novo registro."
      Top             =   3420
      Width           =   795
   End
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   6780
      TabIndex        =   24
      Top             =   3300
      Width           =   2175
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "mov_bancario.frx":833E
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "mov_bancario.frx":98C0
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "mov_bancario.frx":AD32
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "mov_bancario.frx":C22C
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   8160
      Picture         =   "mov_bancario.frx":D726
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Cancela o registro atual."
      Top             =   3420
      Width           =   795
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   7260
      Picture         =   "mov_bancario.frx":EC20
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Confirma o registro atual."
      Top             =   3420
      Width           =   795
   End
End
Attribute VB_Name = "movimento_bancario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lFlagMovimento As Integer
Dim lOpcao As Integer
Dim lConta As String
Dim lData As Date
Dim lNumero As Long
Private ContaBancaria As New cContaBancaria
Private Historico As New cHistorico
Private MovBancario As New cMovimentoBancario
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
    MovBancario.Empresa = g_empresa
    MovBancario.CodigoConta = dtcbo_conta.BoundText
    MovBancario.Data = CDate(msk_data.Text)
    If lOpcao = 1 Then
        MovBancario.NumeroMovimento = 1
    End If
    MovBancario.valor = fValidaValor2(txt_valor.Text)
    MovBancario.DebitoouCredito = cbo_operacao.Text
    MovBancario.NumeroDocumento = txt_documento.Text
    MovBancario.CodigoHistorico = Val(dtcbo_historico.BoundText)
    MovBancario.Complemento = txt_complemento.Text
End Sub
Private Sub AtualTela()
    lConta = MovBancario.CodigoConta
    lData = MovBancario.Data
    lNumero = MovBancario.NumeroMovimento
    dtcbo_conta.BoundText = MovBancario.CodigoConta
    msk_data.Text = Format(MovBancario.Data, "dd/mm/yyyy")
    txt_valor.Text = Format(MovBancario.valor, "###,##0.00")
    cbo_operacao.Text = MovBancario.DebitoouCredito
    txt_documento.Text = MovBancario.NumeroDocumento
    dtcbo_historico.BoundText = MovBancario.CodigoHistorico
    txt_complemento.Text = MovBancario.Complemento
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
Private Sub cbo_operacao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        txt_complemento.SetFocus
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
    txt_valor.SetFocus
End Sub
Private Sub cmd_anterior_Click()
    If MovBancario.LocalizarAnterior Then
        AtualTela
    Else
        MsgBox "Início de Arquivo.", vbInformation, "Atenção!"
        cmd_proximo.SetFocus
    End If
End Sub
Private Sub cmd_cancelar_Click()
    lOpcao = 0
    LimpaTela
    If MovBancario.LocalizarRegistro(g_empresa, lConta, lData, lNumero) Then
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
Private Sub cmd_data_Click()
    g_string = msk_data.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    msk_data.Text = RetiraGString(1)
    txt_valor.SetFocus
    g_string = " "
End Sub
Private Sub cmd_excluir_Click()
    If dtcbo_conta.BoundText <> "" Then
        If (MsgBox("Deseja Realmente Excluir Este Registro?", 4 + 32 + 256, "Exclusão de Registro!")) = 6 Then
            If MovBancario.Excluir(g_empresa, dtcbo_conta.BoundText, CDate(msk_data.Text), lNumero) Then
                LimpaTela
                If MovBancario.LocalizarUltimo(g_empresa, lConta) Then
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
    If lConta <> "" Then
        dtcbo_conta.BoundText = lConta
        msk_data.Text = Format(lData, "dd/mm/yyyy")
        txt_valor.SetFocus
    Else
        dtcbo_conta.SetFocus
    End If
End Sub
Private Sub cmd_novo_KeyPress(KeyAscii As Integer)
    'Crtl+R
    If KeyAscii = 18 Then
        Call MovBancario.RecalculaSaldo(g_empresa, lConta)
    End If
End Sub
Private Sub cmd_ok_Click()

On Error GoTo FileError
    
    If ValidaCampos Then
        AtivaBotoes
        If lOpcao = 1 Then
            AtualTabe
            If MovBancario.Incluir > 0 Then
                lConta = MovBancario.CodigoConta
                lData = MovBancario.Data
                lNumero = MovBancario.NumeroMovimento
            Else
                MsgBox "Não foi possível incluir este registro!", vbInformation, "Erro de Verificação!"
            End If
        ElseIf lOpcao = 2 Then
            AtualTabe
            If MovBancario.Alterar(g_empresa, lConta, lData, lNumero) Then
                lConta = MovBancario.CodigoConta
                lData = MovBancario.Data
                lNumero = MovBancario.NumeroMovimento
            Else
                MsgBox "Não foi possível alterar este registro!", vbInformation, "Erro de Verificação!"
            End If
        End If
        lOpcao = 0
        If MovBancario.LocalizarRegistro(g_empresa, lConta, lData, lNumero) Then
            AtualTela
        Else
            LimpaTela
            MsgBox "Registro não encontrado!", vbInformation, "Erro de Integridade!"
        End If
        cmd_novo.SetFocus
    End If
    Exit Sub
FileError:
    MsgBox Error
    'ErroArquivo tbl_movimento_bancario.Name, "Movimento Bancárioo"
    Exit Sub
End Sub
Private Sub cmd_pesquisa_Click()
    consulta_movimento_bancario.Show 1
    If Len(g_string) > 0 Then
        lConta = RetiraGString(1)
        lData = RetiraGString(2)
        lNumero = RetiraGString(3)
        If MovBancario.LocalizarRegistro(g_empresa, lConta, lData, lNumero) Then
            AtualTela
        Else
            LimpaTela
            MsgBox "Registro não encontrado!", vbInformation, "Erro de Integridade!"
        End If
    End If
End Sub
Private Sub cmd_primeiro_Click()
    If MovBancario.LocalizarPrimeiro(g_empresa, dtcbo_conta.BoundText) Then
        AtualTela
        cmd_proximo.SetFocus
    Else
        LimpaTela
        MsgBox "Não há registros nesta conta.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    If MovBancario.LocalizarProximo Then
        AtualTela
    Else
        MsgBox "Fim de Arquivo.", vbInformation, "Atenção!"
        cmd_anterior.SetFocus
    End If
End Sub
Private Sub Finaliza()
    Set ContaBancaria = Nothing
    Set Historico = Nothing
    Set MovBancario = Nothing
End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub cmd_ultimo_Click()
    If MovBancario.LocalizarUltimo(g_empresa, dtcbo_conta.BoundText) Then
        AtualTela
        cmd_anterior.SetFocus
    Else
        MsgBox "Não há registros nesta conta.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub dtcbo_conta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        If lOpcao > 0 Then
            msk_data.SetFocus
        Else
            If MovBancario.LocalizarUltimo(g_empresa, dtcbo_conta.BoundText) Then
                AtivaBotoes
                AtualTela
                cmd_anterior.SetFocus
            Else
                DesativaBotoes
                cmd_novo.Enabled = True
                cmd_sair.Enabled = True
                LimpaTela
                MsgBox "Não há registros nesta conta.", vbInformation, "Erro de Verificação!"
            End If
        End If
    End If
End Sub
Private Sub dtcbo_conta_LostFocus()
    'If dtcbo_conta.BoundText <> "" Then
    'End If
End Sub
Private Sub dtcbo_historico_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        cbo_operacao.SetFocus
    End If
End Sub

Private Sub dtcbo_historico_LostFocus()
    If dtcbo_historico.BoundText <> "" Then
        If Historico.LocalizarCodigo(g_empresa, dtcbo_historico.BoundText) Then
            cbo_operacao.Text = Historico.Natureza
            txt_complemento.SetFocus
        End If
    End If
End Sub

Private Sub Form_Activate()
    If lFlagMovimento = 0 Then
        DesativaBotoes
        If MovBancario.LocalizarUltimoGeral(g_empresa) Then
            AtualTela
            AtivaBotoes
        Else
            cmd_novo.Enabled = True
            cmd_sair.Enabled = True
        End If
        cmd_novo.SetFocus
    Else
        lFlagMovimento = 0
    End If
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If dtcbo_conta.BoundText = "" Then
        MsgBox "Selecione uma conta.", vbInformation, "Atenção!"
        dtcbo_conta.SetFocus
    ElseIf Not IsDate(msk_data.Text) Then
        MsgBox "Informe uma data válida.", vbInformation, "Atenção!"
        msk_data.SetFocus
    ElseIf Not fValidaValor2(txt_valor.Text) > 0 Then
        MsgBox "Informe o Valor do Movimento.", vbInformation, "Atenção!"
        txt_valor.SetFocus
    ElseIf dtcbo_historico.BoundText = "" Then
        MsgBox "Selecione um histórico.", vbInformation, "Atenção!"
        dtcbo_historico.SetFocus
    ElseIf cbo_operacao.ListIndex = -1 Then
        MsgBox "Selecione uma operação.", vbInformation, "Atenção!"
        cbo_operacao.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub PreencheCboOperacao()
    cbo_operacao.Clear
    cbo_operacao.AddItem "C"
    cbo_operacao.ItemData(cbo_operacao.NewIndex) = 1
    cbo_operacao.AddItem "D"
    cbo_operacao.ItemData(cbo_operacao.NewIndex) = 2
End Sub
Private Sub Form_Deactivate()
    lFlagMovimento = 1
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
    Screen.MousePointer = 0
    CentraForm Me
    
'    adodc_conta.ConnectionString = gConnectionString
'    adodc_conta.RecordSource = "SELECT Codigo, Nome FROM Conta_Bancaria WHERE Empresa = " & g_empresa & " ORDER BY Nome"
'    adodc_conta.Refresh
    Set adodc_conta.Recordset = Conectar.RsConexao("SELECT Codigo, Nome FROM Conta_Bancaria WHERE Empresa = " & g_empresa & " ORDER BY Nome")
'    adodc_historico.ConnectionString = gConnectionString
'    adodc_historico.RecordSource = "SELECT Codigo, Nome FROM Historico WHERE Empresa = " & g_empresa & " ORDER BY Nome"
'    adodc_historico.Refresh
    Set adodc_historico.Recordset = Conectar.RsConexao("SELECT Codigo, Nome FROM Historico WHERE Empresa = " & g_empresa & " ORDER BY Nome")
    PreencheCboOperacao
End Sub
Private Sub Inclui()
    lOpcao = 1
    DesativaBotoes
    cmd_novo.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
End Sub
Private Sub LimpaTela()
    dtcbo_conta.BoundText = ""
    msk_data.Text = "__/__/____"
    txt_valor.Text = ""
    dtcbo_historico.BoundText = ""
    cbo_operacao.ListIndex = -1
    txt_complemento.Text = ""
    txt_documento.Text = ""
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_GotFocus()
    If Not IsDate(msk_data.Text) Then
        msk_data.Text = Format(g_data_def, "dd/mm/yyyy")
    End If
    msk_data.SelStart = 0
    msk_data.SelLength = 2
End Sub
Private Sub msk_data_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_valor.SetFocus
    End If
End Sub
Private Sub txt_complemento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_documento.SetFocus
    End If
End Sub
Private Sub txt_documento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub txt_valor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        dtcbo_historico.SetFocus
    End If
    Call ValidaValor(KeyAscii)
End Sub
Private Sub txt_valor_LostFocus()
    If fValidaValor(txt_valor.Text) > 0 Then
        txt_valor.Text = Format(txt_valor.Text, "###,##0.00")
    End If
End Sub
