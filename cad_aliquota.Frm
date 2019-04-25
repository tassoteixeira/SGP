VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form cadastro_aliquota 
   Caption         =   "Cadastro de Alíquota"
   ClientHeight    =   4935
   ClientLeft      =   630
   ClientTop       =   3315
   ClientWidth     =   6855
   Icon            =   "cad_aliquota.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "cad_aliquota.frx":030A
   ScaleHeight     =   4935
   ScaleWidth      =   6855
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid 
      Height          =   2055
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   3625
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   -2147483633
   End
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   120
      Picture         =   "cad_aliquota.frx":0750
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Cria um novo registro."
      Top             =   3960
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   1020
      Picture         =   "cad_aliquota.frx":1DE2
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Altera o registro atual."
      Top             =   3960
      Width           =   795
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   1920
      Picture         =   "cad_aliquota.frx":32DC
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Exclui o registro atual."
      Top             =   3960
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   2820
      Picture         =   "cad_aliquota.frx":496E
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   3960
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Enabled         =   0   'False
      Height          =   1755
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   6615
      Begin VB.TextBox txt_aliquota 
         Height          =   285
         Left            =   2160
         MaxLength       =   6
         TabIndex        =   8
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txt_nome 
         Height          =   285
         Left            =   2160
         MaxLength       =   30
         TabIndex        =   6
         Top             =   960
         Width           =   4335
      End
      Begin VB.TextBox txt_codigo_fiscal 
         Height          =   285
         Left            =   2160
         MaxLength       =   2
         TabIndex        =   4
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txt_codigo 
         Height          =   285
         Left            =   2160
         MaxLength       =   2
         TabIndex        =   2
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "%"
         Height          =   195
         Left            =   2820
         TabIndex        =   9
         Top             =   1380
         Width           =   195
      End
      Begin VB.Label Label7 
         Caption         =   "&Alíquota do Imposto"
         Height          =   255
         Left            =   420
         TabIndex        =   7
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Descrição do &Imposto"
         Height          =   255
         Left            =   420
         TabIndex        =   5
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Có&digo da Alíquota"
         Height          =   255
         Index           =   2
         Left            =   420
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Código &Fiscal"
         Height          =   255
         Left            =   420
         TabIndex        =   3
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   4560
      TabIndex        =   17
      Top             =   3840
      Width           =   2175
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "cad_aliquota.frx":6000
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "cad_aliquota.frx":7582
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "cad_aliquota.frx":89F4
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "cad_aliquota.frx":9EEE
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   5940
      Picture         =   "cad_aliquota.frx":B3E8
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Cancela o registro atual."
      Top             =   3960
      Width           =   795
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   5040
      Picture         =   "cad_aliquota.frx":C8E2
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Confirma o registro atual."
      Top             =   3960
      Width           =   795
   End
End
Attribute VB_Name = "cadastro_aliquota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lFlagCadastro As Integer
Dim lOpcao As Integer
Dim lCodigo As Integer
Dim lSQl As String
Dim lSerieECF As String
Private rsTabela As New adodb.Recordset
Private Aliquota As New cAliquota
Private Sub AtivaBotoes()
    cmd_novo.Enabled = True
    cmd_alterar.Enabled = True
    cmd_excluir.Enabled = True
    cmd_sair.Enabled = True
    MSFlexGrid.Enabled = True
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
    txt_codigo.Enabled = True
    frm_move.Visible = True
    txt_codigo.Enabled = True
End Sub
Private Sub AtualizaMSFlexGrid()
    Dim i As Integer
    On Error GoTo ErroConsulta
    LimpaMSFlexGrid
    lSQl = "SELECT Codigo, [Codigo Fiscal], Nome, [Aliquota do Imposto]"
    lSQl = lSQl & " FROM Aliquota"
    lSQl = lSQl & " ORDER BY Codigo"
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
            MSFlexGrid.Text = rsTabela("Codigo Fiscal").Value
            MSFlexGrid.Col = 2
            MSFlexGrid.Text = rsTabela("Nome").Value
            MSFlexGrid.Col = 3
            MSFlexGrid.Text = rsTabela("Aliquota do Imposto").Value
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
        Aliquota.Codigo = "" & txt_codigo.Text
    End If
    Aliquota.CodigoFiscal = "" & txt_codigo_fiscal.Text
    Aliquota.Nome = "" & txt_nome.Text
    Aliquota.Aliquota = "" & txt_aliquota.Text
End Sub
Private Sub AtualTela()
    lCodigo = Aliquota.Codigo
    txt_codigo.Text = Aliquota.Codigo
    txt_codigo_fiscal.Text = Aliquota.CodigoFiscal
    txt_nome.Text = Aliquota.Nome
    txt_aliquota.Text = Aliquota.Aliquota
    frm_dados.Enabled = False
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
Private Sub Finaliza()
    Set Aliquota = Nothing
    frm_cadastro.Show
End Sub
Private Sub Inclui()
    lOpcao = 1
    DesativaBotoes
    cmd_novo.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    txt_codigo.Text = 0
    If Aliquota.LocalizarUltimo Then
        txt_codigo.Text = Aliquota.Codigo
    End If
    txt_codigo.Text = Format(Val(txt_codigo.Text) + 1, "00")
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
    If Aliquota.LocalizarAnterior Then
        AtualTela
    Else
        MsgBox "Início de Arquivo.", vbInformation, "Atenção!"
        cmd_proximo.SetFocus
    End If
End Sub
Private Sub cmd_cancelar_Click()
    LimpaTela
    lOpcao = 0
    If Aliquota.LocalizarCodigo(lSerieECF, lCodigo) Then
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
    MSFlexGrid.Text = "Código Fiscal"
    MSFlexGrid.ColWidth(i) = 660
    MSFlexGrid.ColAlignment(i) = flexAlignCenterCenter
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Nome"
    MSFlexGrid.ColWidth(i) = 4000
    MSFlexGrid.ColAlignment(i) = flexAlignGeneral
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Alíquota do Imposto"
    MSFlexGrid.ColWidth(i) = 1000
    MSFlexGrid.ColAlignment(i) = flexAlignGeneral
End Sub
Private Sub LimpaTela()
    txt_codigo.Text = ""
    txt_codigo_fiscal.Text = ""
    txt_nome.Text = ""
    txt_aliquota.Text = ""
End Sub
Private Sub cmd_excluir_Click()
    If Val(txt_codigo) > 0 Then
        If (MsgBox("Deseja excluir este registro?", 4 + 32 + 256, "Exclusão de Registro!")) = 6 Then
            If Aliquota.Excluir(lSerieECF, txt_codigo.Text) Then
                LimpaTela
                If Aliquota.LocalizarUltimo Then
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
    txt_codigo_fiscal.SetFocus
End Sub
Private Sub cmd_ok_Click()
    On Error GoTo FileError
    If ValidaCampos Then
        AtivaBotoes
        If lOpcao = 1 Then
            AtualTabe
            If Aliquota.Incluir Then
                lCodigo = Val(txt_codigo.Text)
            Else
                MsgBox "Não foi possível incluir este registro!", vbInformation, "Erro de Verificação!"
            End If
        ElseIf lOpcao = 2 Then
            AtualTabe
            If Aliquota.Alterar(lSerieECF, lCodigo) Then
            Else
                MsgBox "Não foi possível alterar este registro!", vbInformation, "Erro de Verificação!"
            End If
        End If
        If Aliquota.LocalizarCodigo(lSerieECF, lCodigo) Then
            AtualizaMSFlexGrid
            AtualTela
        End If
        lOpcao = 0
        cmd_novo.SetFocus
    End If
    Exit Sub
FileError:
    MsgBox Error
    'ErroArquivo dta_aliquota.Recordset.Name, "Alíquotaa"
    Exit Sub
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If Not Val(txt_codigo) > 0 Then
        MsgBox "Informe o código da alíquota.", vbInformation, "Atenção!"
        txt_codigo.SetFocus
    ElseIf txt_codigo_fiscal = "" Then
        MsgBox "Informe o código fiscal.", vbInformation, "Atenção!"
        txt_codigo_fiscal.SetFocus
    ElseIf txt_nome = "" Then
        MsgBox "Informe a descrição do imposto.", vbInformation, "Atenção!"
        txt_nome.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_primeiro_Click()
    If Aliquota.LocalizarPrimeiro Then
        AtualTela
        cmd_proximo.SetFocus
    Else
        MsgBox "Não há registro.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    If Aliquota.LocalizarProximo Then
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
    If Aliquota.LocalizarUltimo Then
        AtualTela
        cmd_anterior.SetFocus
    Else
        MsgBox "Não há registro.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub Form_Activate()
    If lFlagCadastro = 0 Then
        DesativaBotoes
        If Aliquota.LocalizarUltimo Then
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
        If Aliquota.LocalizarCodigo(lSerieECF, lCodigo) Then
            AtualTela
        End If
    End If
End Sub
Private Sub txt_aliquota_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub txt_aliquota_LostFocus()
    txt_aliquota = Format(txt_aliquota, "##0.00")
End Sub
Private Sub txt_codigo_fiscal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_nome.SetFocus
    End If
End Sub
Private Sub txt_codigo_GotFocus()
    txt_codigo.SelStart = 0
    txt_codigo.SelLength = Len(txt_codigo)
End Sub
Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_codigo_fiscal.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_codigo_LostFocus()
    If lOpcao = 1 And txt_codigo.Text <> "" Then
        If Aliquota.LocalizarCodigo(lSerieECF, txt_codigo.Text) Then
            MsgBox "Já existe alíquota cadastrada com este código." & Chr(10) & Chr(10) & "Mude o código informado.", vbInformation, "Duplicidade de Registro!"
            txt_codigo.SetFocus
            Exit Sub
        End If
    End If
End Sub
Private Sub txt_nome_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_aliquota.SetFocus
    End If
End Sub
Private Sub txt_nome_LostFocus()
    If lOpcao = 1 And txt_nome.Text <> "" Then
        If Aliquota.LocalizarNome(lSerieECF, txt_nome.Text) Then
            If (MsgBox("Existe outra alíquota cadastrada com o mesmo nome." & Chr(10) & Chr(10) & "Deseja cadastrar assim mesmo?", 4 + 32 + 256, "Duplicidade de Registro!")) = 7 Then
                txt_nome.SetFocus
            End If
        End If
    End If
End Sub
