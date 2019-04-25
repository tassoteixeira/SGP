VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form cadastro_combustivel 
   Caption         =   "Cadastro de Combustíveis"
   ClientHeight    =   5490
   ClientLeft      =   4110
   ClientTop       =   2940
   ClientWidth     =   7470
   Icon            =   "cad_combustivel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "cad_combustivel.frx":030A
   ScaleHeight     =   5490
   ScaleWidth      =   7470
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   120
      Picture         =   "cad_combustivel.frx":0750
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Cria um novo registro."
      Top             =   4560
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   1020
      Picture         =   "cad_combustivel.frx":1DE2
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Altera o registro atual."
      Top             =   4560
      Width           =   795
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   1920
      Picture         =   "cad_combustivel.frx":32DC
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Exclui o registro atual."
      Top             =   4560
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   2820
      Picture         =   "cad_combustivel.frx":496E
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   4560
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   1995
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   7275
      Begin VB.TextBox txt_preco_medio 
         Height          =   285
         Left            =   2520
         MaxLength       =   13
         TabIndex        =   10
         Top             =   1620
         Width           =   1095
      End
      Begin VB.TextBox txt_capacidade_armazenamento 
         Height          =   285
         Left            =   2520
         MaxLength       =   13
         TabIndex        =   8
         Top             =   1260
         Width           =   1095
      End
      Begin VB.TextBox txt_quantidade_estoque 
         Height          =   285
         Left            =   2520
         MaxLength       =   13
         TabIndex        =   6
         Top             =   900
         Width           =   1095
      End
      Begin VB.TextBox txt_codigo 
         Height          =   285
         Left            =   2520
         MaxLength       =   2
         TabIndex        =   2
         Top             =   180
         Width           =   495
      End
      Begin VB.TextBox txt_nome 
         Height          =   285
         Left            =   2520
         MaxLength       =   20
         TabIndex        =   4
         Top             =   540
         Width           =   3675
      End
      Begin VB.Label Label3 
         Caption         =   "&Preço de Venda Médio"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   1620
         Width           =   2355
      End
      Begin VB.Label Label3 
         Caption         =   "Capacidade de &Armazenamento"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   1260
         Width           =   2355
      End
      Begin VB.Label Label3 
         Caption         =   "&Quantidade em Estoque"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   900
         Width           =   2355
      End
      Begin VB.Label Label1 
         Caption         =   "Código do Co&mbustível"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   180
         Width           =   2355
      End
      Begin VB.Label Label2 
         Caption         =   "&Nome do Combustível"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   540
         Width           =   2355
      End
   End
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   5220
      TabIndex        =   18
      Top             =   4440
      Width           =   2175
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "cad_combustivel.frx":6000
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "cad_combustivel.frx":7582
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "cad_combustivel.frx":89F4
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "cad_combustivel.frx":9EEE
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   5700
      Picture         =   "cad_combustivel.frx":B3E8
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Confirma o registro atual."
      Top             =   4560
      Width           =   795
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   6600
      Picture         =   "cad_combustivel.frx":C9F2
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Cancela o registro atual."
      Top             =   4560
      Width           =   795
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid 
      Height          =   2415
      Left            =   120
      TabIndex        =   11
      Top             =   2040
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   4260
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   -2147483633
   End
End
Attribute VB_Name = "cadastro_combustivel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lFlagCadastro As Integer
Dim lOpcao As Integer
Dim lEmpresa As Integer
Dim lCodigo As String
Dim lSQl As String
Private rsTabela As New adodb.Recordset
Private Combustivel As New cCombustivel
Private Sub AtivaBotoes()
    cmd_novo.Enabled = True
    cmd_alterar.Enabled = True
    cmd_excluir.Enabled = True
    cmd_sair.Enabled = True
    MSFlexGrid.Enabled = True
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
    frm_move.Visible = True
    txt_codigo.Enabled = True
End Sub
Private Sub AtualizaMSFlexGrid()
    Dim i As Integer
    On Error GoTo ErroConsulta
    LimpaMSFlexGrid
    lSQl = "SELECT Codigo, Nome, [Quantidade em Estoque], [Capacidade de Armazenamento], [Preco Medio]"
    lSQl = lSQl & " FROM Combustivel"
    lSQl = lSQl & " WHERE Empresa = " & g_empresa
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
            MSFlexGrid.Text = rsTabela("Nome").Value
            MSFlexGrid.Col = 2
            MSFlexGrid.Text = rsTabela("Quantidade em Estoque").Value
            MSFlexGrid.Col = 3
            MSFlexGrid.Text = rsTabela("Capacidade de Armazenamento").Value
            MSFlexGrid.Col = 4
            MSFlexGrid.Text = rsTabela("Preco Medio").Value
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
    Combustivel.Empresa = g_empresa
    Combustivel.Codigo = txt_codigo.Text
    Combustivel.Nome = txt_nome.Text
    Combustivel.QuantidadeEmEstoque = fValidaValor(txt_quantidade_estoque.Text)
    Combustivel.CapacidadeArmazenamento = fValidaValor(txt_capacidade_armazenamento.Text)
    Combustivel.PrecoMedio = fValidaValor(txt_preco_medio.Text)
End Sub
Private Sub AtualTela()
    lCodigo = Combustivel.Codigo
    txt_codigo.Text = Combustivel.Codigo
    txt_nome.Text = Combustivel.Nome
    txt_quantidade_estoque.Text = Format(Combustivel.QuantidadeEmEstoque, "##,###,##0.00")
    txt_capacidade_armazenamento.Text = Format(Combustivel.CapacidadeArmazenamento, "##,###,##0.00")
    txt_preco_medio.Text = Format(Combustivel.PrecoMedio, "###,##0.0000")
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
    Set Combustivel = Nothing
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
    txt_quantidade_estoque.SetFocus
End Sub
Private Sub cmd_anterior_Click()
    If Combustivel.LocalizarAnterior() Then
        AtualTela
    Else
        MsgBox "Início de Arquivo.", vbInformation, "Atenção!"
        cmd_proximo.SetFocus
    End If
End Sub
Private Sub cmd_cancelar_Click()
    LimpaTela
    lOpcao = 0
    If Combustivel.LocalizarCodigo(g_empresa, lCodigo) Then
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
    MSFlexGrid.Cols = 5
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
    MSFlexGrid.ColAlignment(i) = 4
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Nome"
    MSFlexGrid.ColWidth(i) = 2500
    MSFlexGrid.ColAlignment(i) = 1
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Quantidade em Estoque"
    MSFlexGrid.ColWidth(i) = 1200
    MSFlexGrid.ColAlignment(i) = 7
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Capacidade de Armazenamento"
    MSFlexGrid.ColWidth(i) = 1300
    MSFlexGrid.ColAlignment(i) = 7
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Preço de Venda Médio"
    MSFlexGrid.ColWidth(i) = 1200
    MSFlexGrid.ColAlignment(i) = 7
End Sub
Private Sub LimpaTela()
    txt_codigo.Text = ""
    txt_nome.Text = ""
    txt_quantidade_estoque.Text = ""
    txt_capacidade_armazenamento.Text = ""
    txt_preco_medio.Text = ""
End Sub
Private Sub cmd_excluir_Click()
    If txt_codigo.Text <> "" Then
        If (MsgBox("Deseja excluir este registro?", 4 + 32 + 256, "Exclusão de Registro!")) = 6 Then
            If Combustivel.Excluir(g_empresa, txt_codigo.Text) Then
                LimpaTela
                If Combustivel.LocalizarUltimo(g_empresa) Then
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
    txt_codigo.SetFocus
End Sub
Private Sub cmd_ok_Click()
    On Error GoTo FileError
    If ValidaCampos Then
        AtivaBotoes
        If lOpcao = 1 Then
            AtualTabe
            If Combustivel.Incluir Then
                lCodigo = txt_codigo.Text
            Else
                MsgBox "Não foi possível incluir este registro!", vbInformation, "Erro de Verificação!"
            End If
        ElseIf lOpcao = 2 Then
            AtualTabe
            If Not Combustivel.Alterar(g_empresa, txt_codigo.Text) Then
                MsgBox "Não foi possível alterar este registro!", vbInformation, "Erro de Verificação!"
            End If
        End If
        If Combustivel.LocalizarCodigo(g_empresa, lCodigo) Then
            AtualTela
            AtualizaMSFlexGrid
        End If
        lOpcao = 0
        cmd_novo.SetFocus
    End If
    Exit Sub
FileError:
    MsgBox Error
    'ErroArquivo tbl_combustivel.Name, "Combustívelo"
    Exit Sub
End Sub
Function ValidaCampos() As Boolean
    ValidaCampos = False
    If txt_codigo.Text <> "A " And txt_codigo.Text <> "AA" And txt_codigo.Text <> "D " And txt_codigo.Text <> "DA" And txt_codigo.Text <> "G " And txt_codigo.Text <> "GA" Then
        MsgBox "Informe um código de combustível válido.", vbInformation, "Atenção!"
        txt_codigo.SetFocus
    ElseIf Not txt_nome.Text <> "" Then
        MsgBox "Informe o nome do combustível.", vbInformation, "Atenção!"
        txt_nome.SetFocus
    ElseIf Not fValidaValor2(txt_capacidade_armazenamento.Text) > 0 Then
        MsgBox "Informe a capacidade de armazenamento do combustível.", vbInformation, "Atenção!"
        txt_capacidade_armazenamento.SetFocus
    ElseIf Not ValidaPrecoMedio Then
        txt_preco_medio.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Function ValidaPrecoMedio() As Boolean
    Dim xSQL As String
    Dim xPrecoCusto As Currency
    Dim xPrecoVenda As Currency
    ValidaPrecoMedio = False
    xPrecoCusto = 0
    xPrecoVenda = 0
    
    'Prepara SQL
    xSQL = ""
    xSQL = xSQL & "SELECT TOP 1 [Preco de Custo], [Preco de Venda]"
    xSQL = xSQL & "  FROM Bomba"
    xSQL = xSQL & " WHERE Empresa = " & g_empresa
    xSQL = xSQL & "   AND [Tipo de Combustivel] = " & Chr(39) & txt_codigo.Text & Chr(39)
    xSQL = xSQL & " ORDER BY [Tipo de Preco] DESC"
    'Abre RecordSet
    Set rsTabela = New adodb.Recordset
    Set rsTabela = Conectar.RsConexao(xSQL)
    'Verifica movimento
    If rsTabela.RecordCount > 0 Then
        xPrecoCusto = rsTabela("Preco de Custo").Value
        xPrecoVenda = rsTabela("Preco de Venda").Value
    End If
    If rsTabela.State = 1 Then
        rsTabela.Close
    End If
    If fValidaValor(txt_preco_medio.Text) < xPrecoCusto Then
        If (MsgBox("O preço de venda médio não pode ser menor que: " & Format(xPrecoCusto, "###,##0.0000") & "." & Chr(10) & Chr(10) & "Deseja continuar mesmo assim?", vbYesNo + vbDefaultButton2 + vbQuestion, "Preço Está Menor que o de Custo!")) = vbNo Then
            txt_preco_medio.SetFocus
        Else
            ValidaPrecoMedio = True
        End If
    ElseIf fValidaValor(txt_preco_medio.Text) > xPrecoVenda Then
        If (MsgBox("O preço de venda médio não pode ser maior que: " & Format(xPrecoVenda, "###,##0.0000") & "." & Chr(10) & Chr(10) & "Deseja continuar mesmo assim?", vbYesNo + vbDefaultButton2 + vbQuestion, "Preço Está Maior que o de Venda!")) = vbNo Then
            txt_preco_medio.SetFocus
        Else
            ValidaPrecoMedio = True
        End If
    Else
        ValidaPrecoMedio = True
    End If
End Function
Private Sub cmd_primeiro_Click()
    If Combustivel.LocalizarPrimeiro(g_empresa) Then
        AtualTela
        cmd_proximo.SetFocus
    Else
        MsgBox "Não há registro.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    If Combustivel.LocalizarProximo Then
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
    If Combustivel.LocalizarUltimo(g_empresa) Then
        AtualTela
        cmd_anterior.SetFocus
    Else
        MsgBox "Não há registro.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub Form_Activate()
    If g_empresa <> lEmpresa Then
        lFlagCadastro = 0
    End If
    If lFlagCadastro = 0 Then
        lOpcao = 0
        lEmpresa = g_empresa
        DesativaBotoes
        If Combustivel.LocalizarUltimo(g_empresa) Then
            AtivaBotoes
            AtualTela
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
    If KeyCode = vbKeyF2 Then
        KeyCode = 0
        cmd_novo_Click
    ElseIf KeyCode = vbKeyF3 Then
        KeyCode = 0
        cmd_alterar_Click
    ElseIf KeyCode = vbKeyF4 And Shift = 0 Then
        KeyCode = 0
        cmd_excluir_Click
    ElseIf KeyCode = vbKeyF7 Then
        KeyCode = 0
        cmd_primeiro_Click
    ElseIf KeyCode = vbKeyF8 Then
        KeyCode = 0
        cmd_anterior_Click
    ElseIf KeyCode = vbKeyF9 Then
        KeyCode = 0
        cmd_proximo_Click
    ElseIf KeyCode = vbKeyF10 Then
        KeyCode = 0
        cmd_ultimo_Click
    ElseIf KeyCode = vbKeyF11 Then
        KeyCode = 0
        cmd_ok_Click
    ElseIf KeyCode = vbKeyF12 Then
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
        If Combustivel.LocalizarCodigo(g_empresa, lCodigo) Then
            AtualTela
        End If
    End If
End Sub
Private Sub txt_capacidade_armazenamento_GotFocus()
    txt_capacidade_armazenamento.SelStart = 0
    txt_capacidade_armazenamento.SelLength = Len(txt_capacidade_armazenamento.Text)
End Sub
Private Sub txt_capacidade_armazenamento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_preco_medio.SetFocus
    End If
End Sub
Private Sub txt_capacidade_armazenamento_LostFocus()
    txt_capacidade_armazenamento = Format(txt_capacidade_armazenamento, "##,###,##0.00")
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
End Sub
Private Sub txt_codigo_LostFocus()
    If lOpcao = 1 And txt_codigo.Text <> "" Then
        If Combustivel.LocalizarCodigo(g_empresa, txt_codigo.Text) Then
            MsgBox "Já existe combustível cadastrado com este código." & Chr(10) & Chr(10) & "Mude o código informado.", vbInformation, "Duplicidade de Registro!"
            txt_codigo.SetFocus
            Exit Sub
        End If
    End If
End Sub
Private Sub txt_nome_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_quantidade_estoque.SetFocus
    End If
End Sub
Private Sub txt_nome_LostFocus()
    If lOpcao = 1 And txt_nome.Text <> "" Then
        If Combustivel.LocalizarNome(g_empresa, txt_nome.Text) Then
            If (MsgBox("Já existe combustível cadastrado com este nome." & Chr(10) & Chr(10) & "Deseja cadastrar assim mesmo?", 4 + 32 + 256, "Duplicidade de Registro!")) = 7 Then
                txt_nome.SetFocus
            End If
        End If
    End If
End Sub
Private Sub txt_preco_medio_GotFocus()
    txt_preco_medio.SelStart = 0
    txt_preco_medio.SelLength = Len(txt_preco_medio.Text)
End Sub
Private Sub txt_preco_medio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub txt_preco_medio_LostFocus()
    txt_preco_medio.Text = Format(txt_preco_medio.Text, "###,##0.0000")
End Sub
Private Sub txt_quantidade_estoque_GotFocus()
    txt_quantidade_estoque.SelStart = 0
    txt_quantidade_estoque.SelLength = Len(txt_quantidade_estoque.Text)
End Sub
Private Sub txt_quantidade_estoque_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_capacidade_armazenamento.SetFocus
    End If
End Sub
Private Sub txt_quantidade_estoque_LostFocus()
    txt_quantidade_estoque = Format(txt_quantidade_estoque, "##,###,##0.00")
End Sub
