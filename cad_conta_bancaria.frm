VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form cadastro_conta_bancaria 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Cadastro de Conta Bancária"
   ClientHeight    =   2835
   ClientLeft      =   2280
   ClientTop       =   2205
   ClientWidth     =   8175
   ForeColor       =   &H80000008&
   Icon            =   "cad_conta_bancaria.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Bancos"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2835
   ScaleWidth      =   8175
   Begin VB.Frame frm_dados 
      Height          =   1755
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7995
      Begin VB.TextBox txt_agencia 
         Height          =   285
         Left            =   1980
         MaxLength       =   30
         TabIndex        =   9
         Top             =   1320
         Width           =   5220
      End
      Begin VB.TextBox txt_nome 
         Height          =   285
         Left            =   1980
         MaxLength       =   40
         TabIndex        =   4
         Top             =   600
         Width           =   5205
      End
      Begin VB.TextBox txt_codigo 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1980
         MaxLength       =   10
         TabIndex        =   2
         Top             =   240
         Width           =   1305
      End
      Begin VB.CommandButton cmd_mais_banco 
         Caption         =   "+..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7320
         TabIndex        =   7
         Top             =   960
         Width           =   435
      End
      Begin MSAdodcLib.Adodc adodc_banco 
         Height          =   330
         Left            =   4080
         Top             =   960
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
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
         Caption         =   "adodc_banco"
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
      Begin MSDataListLib.DataCombo dtcbo_banco 
         Bindings        =   "cad_conta_bancaria.frx":030A
         Height          =   315
         Left            =   1980
         TabIndex        =   6
         Top             =   960
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   "dtcbo_banco"
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nome da Agência:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Banco:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nome da Conta:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Número da Conta:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   270
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   120
      Picture         =   "cad_conta_bancaria.frx":0324
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Cria um novo registro."
      Top             =   1860
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   1020
      Picture         =   "cad_conta_bancaria.frx":19B6
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Altera o registro atual."
      Top             =   1860
      Width           =   795
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   1920
      Picture         =   "cad_conta_bancaria.frx":2EB0
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Exclui o registro atual."
      Top             =   1860
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   2820
      Picture         =   "cad_conta_bancaria.frx":4542
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   1860
      Width           =   795
   End
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   5940
      TabIndex        =   16
      Top             =   1740
      Width           =   2175
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "cad_conta_bancaria.frx":5BD4
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "cad_conta_bancaria.frx":70CE
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "cad_conta_bancaria.frx":85C8
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "cad_conta_bancaria.frx":9A3A
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   7320
      Picture         =   "cad_conta_bancaria.frx":AFBC
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Cancela o registro atual."
      Top             =   1860
      Width           =   795
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   6420
      Picture         =   "cad_conta_bancaria.frx":C4B6
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Confirma o registro atual."
      Top             =   1860
      Width           =   795
   End
End
Attribute VB_Name = "cadastro_conta_bancaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lFlagCadastro As Integer
Dim lOpcao As Integer
Dim lConta As String
Private Banco As New cBanco
Private ContaBancaria As New cContaBancaria
Private Sub AtivaBotoes()
    cmd_novo.Enabled = True
    cmd_alterar.Enabled = True
    cmd_excluir.Enabled = True
    cmd_sair.Enabled = True
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
    frm_move.Visible = True
    txt_codigo.Enabled = True
End Sub
Private Sub AtualTabe()
    If lOpcao = 1 Then
        ContaBancaria.Codigo = txt_codigo.Text
    End If
    ContaBancaria.Empresa = g_empresa
    ContaBancaria.Nome = txt_nome.Text
    ContaBancaria.Banco = dtcbo_banco.BoundText
    ContaBancaria.Agencia = txt_agencia.Text
End Sub
Private Sub AtualTela()
    lConta = ContaBancaria.Codigo
    txt_codigo.Text = ContaBancaria.Codigo
    txt_nome.Text = ContaBancaria.Nome
    dtcbo_banco.BoundText = ContaBancaria.Banco
    txt_agencia.Text = ContaBancaria.Agencia
    frm_dados.Enabled = False
End Sub
Private Sub DesativaBotoes()
    cmd_novo.Enabled = False
    cmd_alterar.Enabled = False
    cmd_excluir.Enabled = False
    cmd_sair.Enabled = False
    frm_move.Visible = False
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
End Sub
Private Sub cmd_alterar_Click()
    lOpcao = 2
    DesativaBotoes
    cmd_alterar.Visible = True
    cmd_alterar.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    frm_dados.Enabled = True
    txt_codigo.Enabled = False
    txt_nome.SetFocus
End Sub
Private Sub cmd_anterior_Click()
    If ContaBancaria.LocalizarAnterior Then
        AtualTela
    Else
        MsgBox "Início de Arquivo.", vbInformation, "Atenção!"
        cmd_proximo.SetFocus
    End If
End Sub
Private Sub cmd_cancelar_Click()
    LimpaTela
    lOpcao = 0
    If ContaBancaria.LocalizarCodigo(g_empresa, lConta) Then
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
Private Sub cmd_excluir_Click()
    If txt_codigo.Text <> "" Then
        If (MsgBox("Deseja Realmente Excluir Este Registro?", 4 + 32 + 256, "Exclusão de Registro!")) = 6 Then
            If ContaBancaria.Excluir(g_empresa, txt_codigo.Text) Then
                LimpaTela
                If ContaBancaria.LocalizarUltimo(g_empresa) Then
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
Private Sub cmd_mais_banco_Click()
    Dim i As Integer
    Screen.MousePointer = 11
    cadastro_banco.Show 1
    txt_agencia.SetFocus
    dtcbo_banco.BoundText = g_banco
End Sub
Private Sub cmd_novo_Click()
    LimpaTela
    Inclui
    frm_dados.Enabled = True
    txt_codigo.SetFocus
End Sub
Private Sub cmd_ok_Click()
    On Error GoTo FileError
    If ValidaCampos Then
        AtivaBotoes
        If lOpcao = 1 Then
            AtualTabe
            If ContaBancaria.Incluir Then
                lConta = txt_codigo.Text
            Else
                MsgBox "Não foi possível incluir este registro!", vbInformation, "Erro de Verificação!"
            End If
        ElseIf lOpcao = 2 Then
            AtualTabe
            If Not ContaBancaria.Alterar(g_empresa, lConta) Then
                MsgBox "Não foi possível alterar este registro!", vbInformation, "Erro de Verificação!"
            End If
        End If
        lOpcao = 0
        Call ContaBancaria.LocalizarCodigo(g_empresa, lConta)
        AtualTela
        cmd_novo.SetFocus
    End If
    Exit Sub
FileError:
    MsgBox Error
    'ErroArquivo tblConta_bancaria.Name, "Conta Bancáriaa"
    Exit Sub
End Sub
Private Sub cmd_primeiro_Click()
    If ContaBancaria.LocalizarPrimeiro Then
        AtualTela
        cmd_proximo.SetFocus
    Else
        MsgBox "Não há registro.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    If ContaBancaria.LocalizarProximo Then
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
    If ContaBancaria.LocalizarUltimo(g_empresa) Then
        AtualTela
        cmd_anterior.SetFocus
    Else
        MsgBox "Não há registro.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub Finaliza()
    Set Banco = Nothing
    Set ContaBancaria = Nothing
    frm_cadastro.Show
End Sub
Private Sub dtcbo_banco_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        txt_agencia.SetFocus
    End If
End Sub
Private Sub Form_Activate()
    If lFlagCadastro = 0 Then
        DesativaBotoes
        If ContaBancaria.LocalizarUltimo(g_empresa) Then
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
    Screen.MousePointer = 1
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If Not txt_codigo.Text <> "" Then
        MsgBox "Informe código da conta bancária.", vbInformation, "Atenção!"
        txt_codigo.SetFocus
    ElseIf Not txt_nome.Text <> "" Then
        MsgBox "Informe o nome da conta bancária.", vbInformation, "Atenção!"
        txt_nome.SetFocus
    ElseIf dtcbo_banco.BoundText = "" Then
        MsgBox "Selecione o banco.", vbInformation, "Atenção!"
        dtcbo_banco.SetFocus
    ElseIf Not txt_agencia.Text <> "" Then
        MsgBox "Informe o nome da agência.", vbInformation, "Atenção!"
        txt_agencia.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
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
    Screen.MousePointer = 1
    adodc_banco.ConnectionString = Conectar.ConnectionString
    adodc_banco.RecordSource = "SELECT Codigo, Nome FROM Bancos"
    adodc_banco.Refresh
End Sub
Private Sub Inclui()
    lOpcao = 1
    DesativaBotoes
    cmd_novo.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
End Sub
Private Sub LimpaTela()
    txt_codigo.Text = ""
    txt_nome.Text = ""
    dtcbo_banco.BoundText = ""
    txt_agencia.Text = ""
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub txt_agencia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_nome.SetFocus
    End If
End Sub
Private Sub txt_nome_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dtcbo_banco.SetFocus
    End If
End Sub
Private Sub txt_nome_LostFocus()
    If lOpcao = 1 And txt_nome.Text <> "" Then
        If ContaBancaria.LocalizarNome(g_empresa, txt_nome.Text) Then
            If (MsgBox("Já existe conta bancaria cadastrada com este nome." & Chr(10) & Chr(10) & "Deseja cadastrar assim mesmo?", 4 + 32 + 256, "Duplicidade de Registro!")) = 7 Then
                txt_nome.SetFocus
            End If
        End If
    End If
End Sub
