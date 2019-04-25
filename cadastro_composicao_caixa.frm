VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form cadastro_composicao_caixa 
   Caption         =   "Cadastro de Composição de Caixa"
   ClientHeight    =   6795
   ClientLeft      =   1095
   ClientTop       =   1455
   ClientWidth     =   6135
   Icon            =   "cadastro_composicao_caixa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "cadastro_composicao_caixa.frx":030A
   ScaleHeight     =   6795
   ScaleWidth      =   6135
   Begin VB.Frame frm_dados 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   5895
      Begin VB.TextBox txt_configuracao 
         Height          =   285
         Left            =   1740
         MaxLength       =   30
         TabIndex        =   8
         Top             =   1380
         Width           =   795
      End
      Begin VB.TextBox txt_ordem 
         Height          =   285
         Left            =   1740
         MaxLength       =   3
         TabIndex        =   6
         Top             =   1020
         Width           =   435
      End
      Begin VB.TextBox txt_nome 
         Height          =   285
         Left            =   1740
         MaxLength       =   30
         TabIndex        =   4
         Top             =   660
         Width           =   3975
      End
      Begin VB.TextBox txt_codigo 
         Height          =   285
         Left            =   1740
         MaxLength       =   3
         TabIndex        =   2
         Top             =   300
         Width           =   435
      End
      Begin VB.Label Label4 
         Caption         =   "Con&figuração"
         Height          =   255
         Left            =   180
         TabIndex        =   7
         Top             =   1380
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "O&rdem"
         Height          =   255
         Left            =   180
         TabIndex        =   5
         Top             =   1020
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "&Nome"
         Height          =   255
         Left            =   180
         TabIndex        =   3
         Top             =   660
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Có&digo"
         Height          =   255
         Left            =   180
         TabIndex        =   1
         Top             =   300
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   120
      Picture         =   "cadastro_composicao_caixa.frx":0750
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Cria um novo registro."
      Top             =   5880
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   1020
      Picture         =   "cadastro_composicao_caixa.frx":1DE2
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Altera o registro atual."
      Top             =   5880
      Width           =   795
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   1920
      Picture         =   "cadastro_composicao_caixa.frx":32DC
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Exclui o registro atual."
      Top             =   5880
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   2820
      Picture         =   "cadastro_composicao_caixa.frx":496E
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   5880
      Width           =   795
   End
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   3840
      TabIndex        =   16
      Top             =   5760
      Width           =   2175
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "cadastro_composicao_caixa.frx":6000
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "cadastro_composicao_caixa.frx":74FA
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "cadastro_composicao_caixa.frx":89F4
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "cadastro_composicao_caixa.frx":9E66
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
      Left            =   5220
      Picture         =   "cadastro_composicao_caixa.frx":B3E8
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Cancela o registro atual."
      Top             =   5880
      Width           =   795
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   4320
      Picture         =   "cadastro_composicao_caixa.frx":C8E2
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Confirma o registro atual."
      Top             =   5880
      Width           =   795
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid 
      Height          =   3915
      Left            =   120
      TabIndex        =   9
      Top             =   1860
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   6906
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   -2147483633
   End
End
Attribute VB_Name = "cadastro_composicao_caixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lFlagCadastro As Integer
Dim lOpcao As Integer
Dim lCodigo As Integer
Dim lSQl As String
Private rsTabela As New adodb.Recordset
Private ComposicaoCaixa As New cComposicaoCaixa
Private Sub AtualizaMSFlexGrid()
    Dim i As Integer
    On Error GoTo ErroConsulta
    LimpaMSFlexGrid
    lSQl = "SELECT Codigo, Nome, Ordem, Configuracao"
    lSQl = lSQl & " FROM Composicao_Caixa"
    lSQl = lSQl & " ORDER BY Ordem"
    'Abre RecordSet
    Set rsTabela = New adodb.Recordset
    Set rsTabela = Conectar.RsConexao(lSQl)
    'Verifica movimento
    i = 0
    If rsTabela.RecordCount > 0 Then
        Do Until rsTabela.EOF
            MSFlexGrid.Rows = MSFlexGrid.Rows + 1
            i = i + 1
            MSFlexGrid.Row = i
            MSFlexGrid.Col = 0
            MSFlexGrid.Text = rsTabela("Codigo").Value
            MSFlexGrid.Col = 1
            MSFlexGrid.Text = rsTabela("Nome").Value
            MSFlexGrid.Col = 2
            MSFlexGrid.Text = rsTabela("Ordem").Value
            MSFlexGrid.Col = 3
            MSFlexGrid.Text = rsTabela("Configuracao").Value
            rsTabela.MoveNext
        Loop
    End If
    If rsTabela.State = 1 Then
        rsTabela.Close
    End If
    Exit Sub
    
ErroConsulta:
    If Err = 3075 Then
        MsgBox "Condição inválida.", vbExclamation, "Erro de Consulta"
    Else
        MsgBox Error, vbExclamation, "Erro de Consulta"
    End If
    Exit Sub
End Sub
Private Sub AtualTabe()
    If lOpcao = 1 Then
        ComposicaoCaixa.Codigo = "" & Val(txt_codigo.Text)
        ComposicaoCaixa.Ativo = True
    End If
    ComposicaoCaixa.Nome = "" & txt_nome.Text
    ComposicaoCaixa.Ordem = "" & Val(txt_ordem.Text)
    ComposicaoCaixa.Configuracao = "" & txt_configuracao.Text
End Sub
Private Sub AtualTela()
    lCodigo = ComposicaoCaixa.Codigo
    txt_codigo.Text = ComposicaoCaixa.Codigo
    txt_nome.Text = ComposicaoCaixa.Nome
    txt_ordem.Text = ComposicaoCaixa.Ordem
    txt_configuracao.Text = ComposicaoCaixa.Configuracao
    frm_dados.Enabled = False
End Sub
Private Sub Finaliza()
    Set ComposicaoCaixa = Nothing
    frm_cadastro.Show
End Sub
Private Sub Inclui()
    lOpcao = 1
    DesativaBotoes
    cmd_novo.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    txt_codigo.Text = 0
    If ComposicaoCaixa.LocalizarUltimo Then
        txt_codigo.Text = ComposicaoCaixa.Codigo
    End If
    txt_codigo.Text = Format(Val(txt_codigo.Text) + 1, "00")
    txt_ordem.Text = txt_codigo.Text
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
    If ComposicaoCaixa.LocalizarAnterior Then
        AtualTela
    Else
        MsgBox "Início de Arquivo.", vbInformation, "Atenção!"
        cmd_proximo.SetFocus
    End If
End Sub
Private Sub cmd_cancelar_Click()
    LimpaTela
    lOpcao = 0
    If ComposicaoCaixa.LocalizarCodigo(lCodigo) Then
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
Private Sub LimpaMSFlexGrid()
    Dim i As Integer
    MSFlexGrid.WordWrap = True
    MSFlexGrid.Cols = 4
    MSFlexGrid.Rows = 2
    MSFlexGrid.Row = 1
    For i = 0 To (MSFlexGrid.Cols - 1)
        MSFlexGrid.Col = i
        MSFlexGrid.Text = ""
    Next
    MSFlexGrid.RowHeight(0) = 500
    MSFlexGrid.Row = 0
    i = 0
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Código"
    MSFlexGrid.ColWidth(i) = 700
    MSFlexGrid.ColAlignment(i) = flexAlignGeneral
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Nome"
    MSFlexGrid.ColWidth(i) = 3000
    MSFlexGrid.ColAlignment(i) = flexAlignGeneral
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Ordem"
    MSFlexGrid.ColWidth(i) = 700
    MSFlexGrid.ColAlignment(i) = flexAlignGeneral
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Configuração"
    MSFlexGrid.ColWidth(i) = 1100
    MSFlexGrid.ColAlignment(i) = flexAlignGeneral
End Sub
Private Sub LimpaTela()
    txt_codigo.Text = ""
    txt_nome.Text = ""
    txt_ordem.Text = ""
    txt_configuracao.Text = ""
End Sub
Private Sub cmd_excluir_Click()
    If Val(txt_codigo.Text) > 0 Then
        If (MsgBox("Deseja excluir este registro?", 4 + 32 + 256, "Atenção!")) = 6 Then
            If ComposicaoCaixa.Excluir(Val(txt_codigo.Text)) Then
                LimpaTela
                If ComposicaoCaixa.LocalizarUltimo Then
                    AtualTela
                    AtualizaMSFlexGrid
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
    txt_nome.SetFocus
End Sub
Private Sub cmd_ok_Click()
    On Error GoTo FileError
    If ValidaCampos Then
        AtivaBotoes
        If lOpcao = 1 Then
            AtualTabe
            If ComposicaoCaixa.Incluir Then
                lCodigo = Val(txt_codigo.Text)
            Else
                MsgBox "Não foi possível incluir este registro!", vbInformation, "Erro de Verificação!"
            End If
        ElseIf lOpcao = 2 Then
            AtualTabe
            If Not ComposicaoCaixa.Alterar(Val(txt_codigo.Text)) Then
                MsgBox "Não foi possível alterar este registro!", vbInformation, "Erro de Verificação!"
            End If
        End If
        If ComposicaoCaixa.LocalizarCodigo(lCodigo) Then
            AtualTela
            AtualizaMSFlexGrid
        End If
        lOpcao = 0
        cmd_novo.SetFocus
    End If
    Exit Sub
FileError:
    MsgBox Error
    'ErroArquivo dta_ComposicaoCaixa.Recordset.Name, "ComposicaoCaixao"
    Exit Sub
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If Not Val(txt_codigo.Text) > 0 Then
        MsgBox "Informe o código da composição de caixa.", vbInformation, "Atenção!"
        txt_codigo.SetFocus
    ElseIf Not txt_nome.Text > "" Then
        MsgBox "Informe o nome da composição de caixa.", vbInformation, "Atenção!"
        txt_nome.SetFocus
    ElseIf Not Val(txt_ordem.Text) > 0 Then
        MsgBox "Informe a ordem da composição de caixa.", vbInformation, "Atenção!"
        txt_ordem.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_primeiro_Click()
    If ComposicaoCaixa.LocalizarPrimeiro Then
        AtualTela
        cmd_proximo.SetFocus
    Else
        MsgBox "Não há registro.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    If ComposicaoCaixa.LocalizarProximo Then
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
    If ComposicaoCaixa.LocalizarUltimo Then
        AtualTela
        cmd_anterior.SetFocus
    Else
        MsgBox "Não há registro.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub Form_Activate()
    If lFlagCadastro = 0 Then
        DesativaBotoes
        If ComposicaoCaixa.LocalizarUltimo Then
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
Private Sub AtivaBotoes()
    cmd_novo.Enabled = True
    cmd_excluir.Enabled = True
    cmd_alterar.Enabled = True
    cmd_sair.Enabled = True
    MSFlexGrid.Enabled = True
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
    frm_move.Visible = True
    txt_codigo.Enabled = True
End Sub
Private Sub DesativaBotoes()
    cmd_novo.Enabled = False
    cmd_alterar.Enabled = False
    cmd_excluir.Enabled = False
    cmd_sair.Enabled = False
    MSFlexGrid.Enabled = False
    frm_move.Visible = False
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
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
    AtualizaMSFlexGrid
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub MSFlexGrid_RowColChange()
    If lOpcao = 0 And MSFlexGrid.TextMatrix(MSFlexGrid.Row, 0) <> "" Then
        lCodigo = Val(MSFlexGrid.TextMatrix(MSFlexGrid.Row, 0))
        If ComposicaoCaixa.LocalizarCodigo(lCodigo) Then
            AtualTela
        End If
    End If
End Sub
Private Sub txt_codigo_GotFocus()
    txt_codigo.SelStart = 0
    txt_codigo.SelLength = Len(txt_codigo.Text)
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
        If ComposicaoCaixa.LocalizarCodigo(Val(txt_codigo.Text)) Then
            MsgBox "Já existe composição de caixa cadastrada com este código." & Chr(10) & Chr(10) & "Mude o código informado.", vbInformation, "Duplicidade de Registro!"
            txt_codigo.SetFocus
            Exit Sub
        End If
    End If
End Sub
Private Sub txt_configuracao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub txt_nome_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_ordem.SetFocus
    End If
End Sub
Private Sub txt_nome_LostFocus()
    If lOpcao = 1 And txt_nome.Text <> "" Then
        If ComposicaoCaixa.LocalizarNome(txt_nome.Text) Then
            If (MsgBox("Existe outra composição de caixa cadastrada com o mesmo nome." & Chr(10) & Chr(10) & "Deseja cadastrar assim mesmo?", 4 + 32 + 256, "Duplicidade de Registro!")) = 7 Then
                txt_nome.SetFocus
            End If
        End If
    End If
End Sub
Private Sub txt_ordem_GotFocus()
    txt_ordem.SelStart = 0
    txt_ordem.SelLength = Len(txt_ordem.Text)
End Sub
Private Sub txt_ordem_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_configuracao.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
