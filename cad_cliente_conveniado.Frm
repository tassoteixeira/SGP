VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form cadastro_cliente_conveniado 
   Caption         =   "Cadastro de Cliente Conveniado"
   ClientHeight    =   2595
   ClientLeft      =   2100
   ClientTop       =   2565
   ClientWidth     =   7095
   Icon            =   "cad_cliente_conveniado.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "cad_cliente_conveniado.frx":030A
   ScaleHeight     =   2595
   ScaleWidth      =   7095
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3720
      Picture         =   "cad_cliente_conveniado.frx":0750
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   1620
      Width           =   795
   End
   Begin VB.CommandButton cmd_pesquisa 
      Caption         =   "&Pesquisa"
      Height          =   855
      Left            =   2820
      Picture         =   "cad_cliente_conveniado.frx":1DE2
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Pesquisa um registro específico."
      Top             =   1620
      Width           =   795
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   1920
      Picture         =   "cad_cliente_conveniado.frx":3254
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Exclui o registro atual."
      Top             =   1620
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   1020
      Picture         =   "cad_cliente_conveniado.frx":48E6
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Altera o registro atual."
      Top             =   1620
      Width           =   795
   End
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   120
      Picture         =   "cad_cliente_conveniado.frx":5DE0
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Cria um novo registro."
      Top             =   1620
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Enabled         =   0   'False
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   6855
      Begin VB.TextBox txt_codigo 
         Height          =   285
         Left            =   1800
         MaxLength       =   6
         TabIndex        =   4
         Top             =   600
         Width           =   675
      End
      Begin VB.TextBox txt_nome 
         Height          =   285
         Left            =   1800
         MaxLength       =   40
         TabIndex        =   6
         Top             =   960
         Width           =   4935
      End
      Begin MSAdodcLib.Adodc adodc_convenio 
         Height          =   330
         Left            =   3780
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
         Caption         =   "adodc_convenio"
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
      Begin MSDataListLib.DataCombo dtcbo_convenio 
         Bindings        =   "cad_cliente_conveniado.frx":7472
         Height          =   315
         Left            =   1800
         TabIndex        =   2
         Top             =   180
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   "dtcbo_convenio"
      End
      Begin VB.Label Label3 
         Caption         =   "Convenio"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1635
      End
      Begin VB.Label Label3 
         Caption         =   "Código do conveniado"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1635
      End
      Begin VB.Label Label4 
         Caption         =   "Nome do conveniado"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1635
      End
   End
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   4800
      TabIndex        =   14
      Top             =   1500
      Width           =   2175
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "cad_cliente_conveniado.frx":748F
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "cad_cliente_conveniado.frx":8A11
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "cad_cliente_conveniado.frx":9E83
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "cad_cliente_conveniado.frx":B37D
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   6180
      Picture         =   "cad_cliente_conveniado.frx":C877
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Cancela o registro atual."
      Top             =   1620
      Width           =   795
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   5280
      Picture         =   "cad_cliente_conveniado.frx":DD71
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Confirma o registro atual."
      Top             =   1620
      Width           =   795
   End
End
Attribute VB_Name = "cadastro_cliente_conveniado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lFlagCadastro As Integer
Dim lOpcao As Integer
Dim lCodigoConvenio As Integer
Dim lCodigoConveniado As Long
Private ClienteConveniado As New cClienteConveniado
Private Sub AtivaBotoes()
    cmd_novo.Enabled = True
    cmd_alterar.Enabled = True
    cmd_excluir.Enabled = True
    cmd_pesquisa.Enabled = True
    cmd_sair.Enabled = True
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
    frm_move.Visible = True
    txt_codigo.Enabled = True
End Sub
Private Sub AtualTabe()
    lCodigoConvenio = Val(dtcbo_convenio.BoundText)
    lCodigoConveniado = Val(txt_codigo.Text)
    If lOpcao = 1 Then
        ClienteConveniado.CodigoConveniado = txt_codigo.Text
    End If
    ClienteConveniado.CodigoConvenio = Val(dtcbo_convenio.BoundText)
    ClienteConveniado.CodigoConveniado = Val(txt_codigo.Text)
    ClienteConveniado.Nome = txt_nome.Text
End Sub
Private Sub AtualTela()
    lCodigoConvenio = ClienteConveniado.CodigoConvenio
    lCodigoConveniado = ClienteConveniado.CodigoConveniado
    dtcbo_convenio.BoundText = ClienteConveniado.CodigoConvenio
    txt_codigo.Text = ClienteConveniado.CodigoConveniado
    txt_nome.Text = ClienteConveniado.Nome
    frm_dados.Enabled = False
End Sub
Private Sub CriaNovoCodigo()
    txt_codigo.Text = 1
    txt_codigo.Text = ClienteConveniado.LocalizarProximoCodigo(Val(dtcbo_convenio.BoundText))
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
Private Sub Finaliza()
    Set ClienteConveniado = Nothing
    frm_cadastro.Show
End Sub
Private Sub Inclui()
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
    frm_dados.Enabled = True
    txt_codigo.Enabled = False
    txt_nome.SetFocus
End Sub
Private Sub cmd_anterior_Click()
    If ClienteConveniado.LocalizarAnterior Then
        AtualTela
    Else
        MsgBox "Início de Arquivo.", vbInformation, "Atenção!"
        cmd_proximo.SetFocus
    End If
End Sub
Private Sub cmd_cancelar_Click()
    LimpaTela
    lOpcao = 0
    If ClienteConveniado.LocalizarCodigo(lCodigoConvenio, lCodigoConveniado) Then
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
    dtcbo_convenio.BoundText = ""
    txt_codigo = ""
    txt_nome = ""
End Sub
Private Sub cmd_excluir_Click()
    If txt_codigo.Text <> "" Then
        If (MsgBox("Deseja excluir este registro?", 4 + 32 + 256, "Exclusão de Registro!")) = 6 Then
            If ClienteConveniado.Excluir(Val(dtcbo_convenio.BoundText), CLng(txt_codigo.Text)) Then
                LimpaTela
                If ClienteConveniado.LocalizarUltimo Then
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
    dtcbo_convenio.SetFocus
End Sub
Private Sub cmd_ok_Click()
    On Error GoTo FileError
    If ValidaCampos Then
        AtivaBotoes
        If lOpcao = 1 Then
            AtualTabe
            If ClienteConveniado.Incluir Then
                lCodigoConvenio = Val(dtcbo_convenio.BoundText)
                lCodigoConveniado = Val(txt_codigo.Text)
            Else
                MsgBox "Não foi possível incluir este registro!", vbInformation, "Erro de Verificação!"
            End If
        ElseIf lOpcao = 2 Then
            AtualTabe
            If Not ClienteConveniado.Alterar(lCodigoConvenio, lCodigoConveniado) Then
                MsgBox "Não foi possível alterar este registro!", vbInformation, "Erro de Verificação!"
            End If
        End If
        lOpcao = 0
        Call ClienteConveniado.LocalizarCodigo(lCodigoConvenio, lCodigoConveniado)
        AtualTela
        cmd_novo.SetFocus
    End If
    Exit Sub
FileError:
    MsgBox Error
    'ErroArquivo tbl_cliente_conveniado.Name, "Cliente Conveniadoo"
    Exit Sub
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If Val(dtcbo_convenio.BoundText) = 0 Then
        MsgBox "Informe o convênio.", vbInformation, "Atenção!"
        dtcbo_convenio.SetFocus
    ElseIf Not Val(txt_codigo) > 0 Then
        MsgBox "Informe o código do conveniado.", vbInformation, "Atenção!"
        txt_codigo.SetFocus
    ElseIf txt_nome = "" Then
        MsgBox "Informe a nome do conveniado.", vbInformation, "Atenção!"
        txt_nome.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_pesquisa_Click()
    consulta_cliente_conveniado.Show 1
    If Len(g_string) > 0 Then
        lCodigoConvenio = RetiraGString(1)
        lCodigoConveniado = RetiraGString(2)
        If ClienteConveniado.LocalizarCodigo(lCodigoConvenio, lCodigoConveniado) Then
            AtualTela
        End If
    End If
End Sub
Private Sub cmd_primeiro_Click()
    If ClienteConveniado.LocalizarPrimeiro Then
        AtualTela
        cmd_proximo.SetFocus
    Else
        MsgBox "Não há registro.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    If ClienteConveniado.LocalizarProximo Then
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
    If ClienteConveniado.LocalizarUltimo Then
        AtualTela
        cmd_anterior.SetFocus
    Else
        MsgBox "Não há registro.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub dtcbo_convenio_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If lOpcao = 1 Then
            txt_codigo.SetFocus
        End If
    End If
End Sub
Private Sub dtcbo_convenio_LostFocus()
    If lOpcao = 1 And Val(dtcbo_convenio.BoundText) > 0 Then
        CriaNovoCodigo
    End If
End Sub
Private Sub Form_Activate()
    If lFlagCadastro = 0 Then
        DesativaBotoes
        If ClienteConveniado.LocalizarUltimo Then
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
    adodc_convenio.ConnectionString = Conectar.ConnectionString
    adodc_convenio.RecordSource = "SELECT Codigo, Nome FROM Convenio"
    adodc_convenio.Refresh
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub

Private Sub txt_codigo_GotFocus()
    If lOpcao = 1 And txt_codigo = "" Then
        CriaNovoCodigo
    End If
    txt_codigo.SelStart = 0
    txt_codigo.SelLength = Len(txt_codigo)
End Sub
Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_nome.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_codigo_LostFocus()
    If lOpcao = 1 And txt_codigo.Text <> "" Then
        If ClienteConveniado.LocalizarCodigo(Val(dtcbo_convenio.BoundText), CLng(txt_codigo.Text)) Then
            MsgBox "Já existe cliente conveniado cadastrado com este código." & Chr(10) & Chr(10) & "Mude o código informado.", vbInformation, "Duplicidade de Registro!"
            txt_codigo.SetFocus
        End If
    End If
End Sub
Private Sub txt_nome_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub txt_nome_LostFocus()
    If lOpcao = 1 And txt_nome <> "" Then
        If ClienteConveniado.LocalizarNome(Val(dtcbo_convenio.BoundText), txt_nome.Text) Then
            If (MsgBox("Já existe cliente conveniado cadastrado com este nome." & Chr(10) & Chr(10) & "Deseja cadastrar assim mesmo?", 4 + 32 + 256, "Duplicidade de Registro!")) = 7 Then
                txt_nome.SetFocus
            End If
        End If
    End If
End Sub
