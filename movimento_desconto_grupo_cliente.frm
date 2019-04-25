VERSION 5.00
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form movimento_desconto_grupo_cliente 
   Caption         =   "Movimento de Desconto de Grupo de Cliente"
   ClientHeight    =   5775
   ClientLeft      =   165
   ClientTop       =   285
   ClientWidth     =   7875
   Icon            =   "movimento_desconto_grupo_cliente.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "movimento_desconto_grupo_cliente.frx":030A
   ScaleHeight     =   5775
   ScaleWidth      =   7875
   Begin MSGrid.Grid grid_dados 
      Height          =   2535
      Left            =   120
      TabIndex        =   18
      Top             =   2220
      Width           =   7635
      _Version        =   65536
      _ExtentX        =   13467
      _ExtentY        =   4471
      _StockProps     =   77
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Cols            =   9
      FixedCols       =   0
   End
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   120
      Picture         =   "movimento_desconto_grupo_cliente.frx":0750
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Cria um novo registro."
      Top             =   4860
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   1020
      Picture         =   "movimento_desconto_grupo_cliente.frx":1DE2
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Altera o registro atual."
      Top             =   4860
      Width           =   795
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   1920
      Picture         =   "movimento_desconto_grupo_cliente.frx":32DC
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Exclui o registro atual."
      Top             =   4860
      Width           =   795
   End
   Begin VB.CommandButton cmd_pesquisa 
      Caption         =   "&Pesquisa"
      Height          =   855
      Left            =   2820
      Picture         =   "movimento_desconto_grupo_cliente.frx":496E
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Pesquisa um registro específico."
      Top             =   4860
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3720
      Picture         =   "movimento_desconto_grupo_cliente.frx":5DE0
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   4860
      Width           =   795
   End
   Begin VB.Frame frmDados 
      Enabled         =   0   'False
      Height          =   2115
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   7635
      Begin VB.TextBox txtPrecoECF 
         Height          =   285
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   17
         Top             =   1740
         Width           =   1095
      End
      Begin MSAdodcLib.Adodc adodcProduto 
         Height          =   330
         Left            =   4200
         Top             =   660
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
         Caption         =   "adodcProduto"
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
      Begin VB.TextBox txtPrecoFixo 
         Height          =   285
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   13
         Top             =   1380
         Width           =   1095
      End
      Begin VB.CheckBox chkAcrescimo 
         Height          =   255
         Left            =   6420
         TabIndex        =   15
         Top             =   1380
         Width           =   1035
      End
      Begin VB.TextBox txt_grupo_cliente 
         Height          =   285
         Left            =   2040
         MaxLength       =   5
         TabIndex        =   2
         Top             =   180
         Width           =   555
      End
      Begin VB.TextBox txt_percentual_descontar 
         Height          =   285
         Left            =   2040
         MaxLength       =   6
         TabIndex        =   8
         Top             =   1020
         Width           =   615
      End
      Begin VB.TextBox txt_produto 
         Height          =   285
         Left            =   2040
         MaxLength       =   6
         TabIndex        =   5
         Top             =   660
         Width           =   795
      End
      Begin VB.TextBox txt_valor_descontar 
         Height          =   285
         Left            =   6420
         MaxLength       =   10
         TabIndex        =   11
         Top             =   1020
         Width           =   1095
      End
      Begin MSAdodcLib.Adodc adodcGrupoCliente 
         Height          =   330
         Left            =   4140
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
         Caption         =   "adodcGrupoCliente"
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
      Begin MSDataListLib.DataCombo dtcboGrupoCliente 
         Bindings        =   "movimento_desconto_grupo_cliente.frx":7472
         Height          =   315
         Left            =   2640
         TabIndex        =   3
         Top             =   180
         Width           =   4875
         _ExtentX        =   8599
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   "dtcboGrupoCliente"
      End
      Begin MSDataListLib.DataCombo dtcboProduto 
         Bindings        =   "movimento_desconto_grupo_cliente.frx":7492
         Height          =   315
         Left            =   2880
         TabIndex        =   6
         Top             =   660
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   "dtcboCliente"
      End
      Begin VB.Label Label3 
         Caption         =   "Preço para EC&F"
         Height          =   315
         Index           =   7
         Left            =   120
         TabIndex        =   16
         Top             =   1740
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "&Preço Fixo"
         Height          =   315
         Index           =   6
         Left            =   120
         TabIndex        =   12
         Top             =   1380
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "&Inverte operação"
         Height          =   315
         Index           =   4
         Left            =   4500
         TabIndex        =   14
         Top             =   1380
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "%"
         Height          =   315
         Index           =   2
         Left            =   2760
         TabIndex        =   9
         Top             =   1080
         Width           =   195
      End
      Begin VB.Line Line1 
         X1              =   20
         X2              =   7650
         Y1              =   590
         Y2              =   590
      End
      Begin VB.Label Label3 
         Caption         =   "&Grupo de Cliente"
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   1
         Top             =   180
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "&Percentual à descontar"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   1020
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "P&roduto"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   660
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "&Valor à descontar"
         Height          =   315
         Index           =   5
         Left            =   4500
         TabIndex        =   10
         Top             =   1020
         Width           =   1815
      End
   End
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   5580
      TabIndex        =   26
      Top             =   4740
      Width           =   2175
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "movimento_desconto_grupo_cliente.frx":74AD
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "movimento_desconto_grupo_cliente.frx":89A7
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "movimento_desconto_grupo_cliente.frx":9EA1
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "movimento_desconto_grupo_cliente.frx":B313
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   6060
      Picture         =   "movimento_desconto_grupo_cliente.frx":C895
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Confirma o registro atual."
      Top             =   4860
      Width           =   795
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   6960
      Picture         =   "movimento_desconto_grupo_cliente.frx":DE9F
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Cancela o registro atual."
      Top             =   4860
      Width           =   795
   End
End
Attribute VB_Name = "movimento_desconto_grupo_cliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lFlagMovimento As Integer
Dim lOpcao As String
Dim lCodigoGrupoCliente As Long
Dim lCodigoProduto As Long
Dim lSQL As String
Dim rsClientes As New adodb.Recordset

Private Cliente As New cCliente
Private GrupoCliente As New cGrupoCliente
Private MovDescontoGrupoCliente As New cMovDescontoGrupoCliente
Private Produto As New cProduto
Private Sub AdcionaDadosGrid()
    Dim x_i As Integer
'    If grid_dados.Rows > 2 Then
'        For x_i = 2 To grid_dados.Rows
'            grid_dados.Row = x_i - 1
'            grid_dados.Col = 2
'            If Val(grid_dados.Text) = !Codigo Then
'                x_flag = False
'                Exit For
'            End If
'        Next
'    End If
    grid_dados.Row = grid_dados.Rows - 1
    grid_dados.Col = 0
    grid_dados.Text = MovDescontoGrupoCliente.CodigoGrupoCliente
    grid_dados.Col = 1
    If GrupoCliente.LocalizarCodigo(MovDescontoGrupoCliente.CodigoGrupoCliente) Then
        grid_dados.Text = GrupoCliente.Nome
    Else
        grid_dados.Text = "** Não Cadastrado **"
    End If
    grid_dados.Col = 2
    grid_dados.Text = MovDescontoGrupoCliente.CodigoProduto
    grid_dados.Col = 3
    If Produto.LocalizarCodigo(MovDescontoGrupoCliente.CodigoProduto) Then
        grid_dados.Text = Produto.Nome
    Else
        grid_dados.Text = "** Não Cadastrado **"
    End If
    grid_dados.Rows = grid_dados.Rows + 1
    grid_dados.Col = 4
    grid_dados.Text = Format(MovDescontoGrupoCliente.PercentualaDescontar, "##0.00")
    grid_dados.Col = 5
    grid_dados.Text = Format(MovDescontoGrupoCliente.ValoraDescontar, "###,##0.0000")
    grid_dados.Col = 6
    grid_dados.Text = Format(MovDescontoGrupoCliente.PrecoFixo, "###,##0.0000")
    grid_dados.Col = 7
    If MovDescontoGrupoCliente.Desconto = True Then
        grid_dados.Text = "Desconto"
    Else
        grid_dados.Text = "Acréscimo"
    End If
    grid_dados.Col = 8
    grid_dados.Text = Format(MovDescontoGrupoCliente.PrecoParaECF, "###,##0.0000")
End Sub
Private Sub AlteraClientesDoMesmoGrupo(ByVal pCodigoGrupoCliente As Integer, ByVal pCodigoProduto As Integer)
    lSQL = ""
    'lSQL = lSQL & "SELECT Codigo"
    lSQL = lSQL & "SELECT Codigo"
    'lSQL = lSQL & "  FROM Cliente"
    lSQL = lSQL & "  FROM GrupoCliente"
    'lSQL = lSQL & " WHERE [Codigo do Grupo de Cliente] = " & pCodigoGrupoCliente
    lSQL = lSQL & " WHERE Codigo = " & pCodigoGrupoCliente
    'lSQL = lSQL & "   AND Codigo <> " & lCodigoGrupoCliente
    lSQL = lSQL & "   AND Codigo <> " & lCodigoGrupoCliente
    lSQL = lSQL & " ORDER BY Codigo"
    Set rsClientes = Conectar.RsConexao(lSQL)
    
    If rsClientes.RecordCount > 0 Then
        If (MsgBox("Existe " & rsClientes.RecordCount & " Cliente(s) no mesmo Grupo de Cliente." & vbCrLf & "Deseja alterar para este(s) outro(s) cliente(s)?", vbYesNo + vbQuestion + vbDefaultButton1, "Alteração de Preço Personalizado!")) = vbYes Then
            Call GravaAuditoria(1, Me.name, 10, "Será alterado " & rsClientes.RecordCount & " Preço de Cliente(s)")
            rsClientes.MoveFirst
            Do Until rsClientes.EOF
                If MovDescontoGrupoCliente.LocalizarCodigo(rsClientes("Codigo").Value, pCodigoProduto) Then
                    AtualizaTabela2
                    If MovDescontoGrupoCliente.Alterar(rsClientes("Codigo").Value, pCodigoProduto) Then
                        Call GravaAuditoria(1, Me.name, 10, "Alterado o preço pers. Cliente:" & rsClientes("Codigo").Value)
                    Else
                        Call GravaAuditoria(1, Me.name, 10, "ERRO ao alterar o preço pers. Cliente:" & rsClientes("Codigo").Value)
                        MsgBox "Não foi possível alterar o Preço Personalizado!" & vbCrLf & "Codigo do Cliente: " & rsClientes("Codigo").Value, vbInformation, "Erro de Integridade!"
                    End If
                Else
                    Call GravaAuditoria(1, Me.name, 10, "ERRO Preço pers. inexistente. Cliente:" & rsClientes("Codigo").Value)
                    MsgBox "Registro de Preço Personalizado Inexistente!" & vbCrLf & "Codigo do Cliente: " & rsClientes("Codigo").Value, vbInformation, "Erro de Integridade!"
                End If
                rsClientes.MoveNext
            Loop
        End If
    End If
    
    'Posiciona Registro atual
    If Not MovDescontoGrupoCliente.LocalizarCodigo(lCodigoGrupoCliente, pCodigoProduto) Then
        MsgBox "Registro ATUAL de Preço Personalizado Inexistente!", vbInformation, "Erro de Integridade!"
    End If
End Sub
Private Sub AtivaBotoes()
    cmd_novo.Enabled = True
    cmd_alterar.Enabled = True
    cmd_excluir.Enabled = True
    If g_nivel_acesso > 4 Then
        If g_empresa < g_cfg_empresa_i Or g_empresa > g_cfg_empresa_f Then
            cmd_novo.Enabled = False
            cmd_alterar.Enabled = False
            cmd_excluir.Enabled = False
        End If
    End If
    cmd_pesquisa.Enabled = True
    cmd_sair.Enabled = True
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
    frm_move.Visible = True
    
    'temporariamente, ate ser criada a pesquisa.
    cmd_pesquisa.Enabled = False
End Sub
Private Sub Inclui()
    lOpcao = 1
    DesativaBotoes
    cmd_novo.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
End Sub
Private Sub AtualTabe()
    MovDescontoGrupoCliente.CodigoGrupoCliente = CLng(dtcboGrupoCliente.BoundText)
    MovDescontoGrupoCliente.CodigoProduto = CLng(dtcboProduto.BoundText)
    AtualizaTabela2
End Sub
Private Sub AtualizaTabela2()
    MovDescontoGrupoCliente.PercentualaDescontar = fValidaValor(txt_percentual_descontar.Text)
    MovDescontoGrupoCliente.ValoraDescontar = fValidaValor(txt_valor_descontar.Text)
    If chkAcrescimo.Value = 0 Then
        MovDescontoGrupoCliente.Desconto = True
    Else
        MovDescontoGrupoCliente.Desconto = False
    End If
    MovDescontoGrupoCliente.PrecoFixo = fValidaValor(txtPrecoFixo.Text)
    MovDescontoGrupoCliente.PrecoParaECF = fValidaValor(txtPrecoECF.Text)
End Sub
Private Sub AtualTela()
    lCodigoGrupoCliente = MovDescontoGrupoCliente.CodigoGrupoCliente
    lCodigoProduto = MovDescontoGrupoCliente.CodigoProduto
    
    txt_grupo_cliente.Text = MovDescontoGrupoCliente.CodigoGrupoCliente
    txt_produto.Text = MovDescontoGrupoCliente.CodigoProduto
    dtcboGrupoCliente.BoundText = MovDescontoGrupoCliente.CodigoGrupoCliente
    dtcboProduto.BoundText = MovDescontoGrupoCliente.CodigoProduto
    txt_percentual_descontar.Text = Format(MovDescontoGrupoCliente.PercentualaDescontar, "##0.00")
    txt_valor_descontar.Text = Format(MovDescontoGrupoCliente.ValoraDescontar, "###,##0.0000")
    If MovDescontoGrupoCliente.Desconto = True Then
        chkAcrescimo.Value = 0
    Else
        chkAcrescimo.Value = 1
    End If
    txtPrecoFixo.Text = Format(MovDescontoGrupoCliente.PrecoFixo, "###,##0.0000")
    txtPrecoECF.Text = Format(MovDescontoGrupoCliente.PrecoParaECF, "###,##0.0000")
    frmDados.Enabled = False
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
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set Cliente = Nothing
    Set MovDescontoGrupoCliente = Nothing
    Set Produto = Nothing
    Set rsClientes = Nothing
End Sub
Private Sub PesquisaPrecoPersonalizado()
    Dim xCodigoGrupoCliente As Long
    Dim xCodigoProduto As Long
    lCodigoGrupoCliente = MovDescontoGrupoCliente.CodigoGrupoCliente
    lCodigoProduto = MovDescontoGrupoCliente.CodigoProduto
    MontaGrid
    xCodigoGrupoCliente = MovDescontoGrupoCliente.CodigoGrupoCliente
    xCodigoProduto = 0
    Do Until MovDescontoGrupoCliente.LocalizarCliente(xCodigoGrupoCliente, xCodigoProduto) = False
        AdcionaDadosGrid
        xCodigoProduto = MovDescontoGrupoCliente.CodigoProduto
    Loop
    Call MovDescontoGrupoCliente.LocalizarCodigo(lCodigoGrupoCliente, lCodigoProduto)
    grid_dados.Row = grid_dados.Rows - 1
    grid_dados.Col = 0
End Sub
Private Function PreparaLogTabela() As String
    Dim xStringLog As String

    xStringLog = "Cli:" & MovDescontoGrupoCliente.CodigoGrupoCliente & " Prod:" & MovDescontoGrupoCliente.CodigoProduto
    If MovDescontoGrupoCliente.Desconto = True Then
        xStringLog = xStringLog & " Desc."
    Else
        xStringLog = xStringLog & " Acresc."
    End If
    If MovDescontoGrupoCliente.PercentualaDescontar > 0 Then
        xStringLog = xStringLog & Format(MovDescontoGrupoCliente.PercentualaDescontar, "##0.00") & "%"
    ElseIf MovDescontoGrupoCliente.ValoraDescontar > 0 Then
        xStringLog = xStringLog & "R$ " & Format(MovDescontoGrupoCliente.ValoraDescontar, "###,##0.0000")
    End If
    If MovDescontoGrupoCliente.PrecoFixo > 0 Then
        xStringLog = xStringLog & " Fixo=" & Format(MovDescontoGrupoCliente.PrecoFixo, "###,##0.0000")
    End If
    If MovDescontoGrupoCliente.PrecoParaECF > 0 Then
        xStringLog = xStringLog & " ECF=" & Format(MovDescontoGrupoCliente.PrecoParaECF, "###,##0.0000")
    End If
    PreparaLogTabela = xStringLog
End Function


Private Sub cmd_alterar_Click()
    Call GravaAuditoria(1, Me.name, 3, "")
    lOpcao = 2
    DesativaBotoes
    cmd_alterar.Visible = True
    cmd_alterar.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    frmDados.Enabled = True
    txt_percentual_descontar.SetFocus
End Sub
Private Sub cmd_anterior_Click()
    Call GravaAuditoria(1, Me.name, 13, "")
    If MovDescontoGrupoCliente.LocalizarAnterior Then
        PesquisaPrecoPersonalizado
        AtualTela
    Else
        MsgBox "Início de Arquivo.", vbInformation, "Atenção!"
        cmd_proximo.SetFocus
    End If
End Sub
Private Sub cmd_cancelar_Click()
    Call GravaAuditoria(1, Me.name, 9, "")
    LimpaTela
    If MovDescontoGrupoCliente.LocalizarCodigo(lCodigoGrupoCliente, lCodigoProduto) Then
        AtivaBotoes
        PesquisaPrecoPersonalizado
        AtualTela
        If cmd_alterar.Enabled Then
            cmd_alterar.SetFocus
        Else
            cmd_novo.SetFocus
        End If
    Else
        DesativaBotoes
        cmd_novo.Enabled = True
        cmd_sair.Enabled = True
        cmd_novo.SetFocus
    End If
    lOpcao = 0
End Sub
Sub LimpaGrid()
    Do Until grid_dados.Rows = 2
        grid_dados.Row = grid_dados.Rows - 1
        grid_dados.RemoveItem grid_dados.Row
    Loop
    grid_dados.Row = 1
    grid_dados.Col = 0
    grid_dados.Text = ""
    grid_dados.Col = 1
    grid_dados.Text = ""
    grid_dados.Col = 2
    grid_dados.Text = ""
    grid_dados.Col = 3
    grid_dados.Text = ""
    grid_dados.Col = 4
    grid_dados.Text = ""
    grid_dados.Col = 5
    grid_dados.Text = ""
    grid_dados.Col = 6
    grid_dados.Text = ""
    grid_dados.Col = 7
    grid_dados.Text = ""
    grid_dados.Col = 8
    grid_dados.Text = ""
End Sub
Private Sub LimpaTela()
    txt_grupo_cliente.Text = ""
    dtcboGrupoCliente.BoundText = ""
    txt_produto.Text = ""
    dtcboProduto.BoundText = ""
    txt_valor_descontar.Text = ""
    txt_percentual_descontar.Text = ""
    txtPrecoFixo.Text = ""
    txtPrecoECF.Text = ""
    chkAcrescimo.Value = 0
End Sub
Private Sub cmd_excluir_Click()
    Call GravaAuditoria(1, Me.name, 4, "")
    If txt_grupo_cliente > 0 Then
        If (MsgBox("Deseja excluir este registro?", 4 + 32 + 256, "Exclusão de Registro!")) = vbYes Then
            Call GravaAuditoria(1, Me.name, 10, PreparaLogTabela)
            lOpcao = 3
            If MovDescontoGrupoCliente.Excluir(CLng(dtcboGrupoCliente.BoundText), CLng(dtcboProduto.BoundText)) Then
                If Not MovDescontoGrupoCliente.LocalizarUltimo() Then
                    DesativaBotoes
                    cmd_novo.Enabled = True
                    cmd_sair.Enabled = True
                    cmd_novo.SetFocus
                End If
                lOpcao = 0
                PesquisaPrecoPersonalizado
                AtualTela
            End If
        End If
    End If
End Sub
Private Sub cmd_novo_Click()
    Call GravaAuditoria(1, Me.name, 2, "")
    LimpaTela
    Inclui
    frmDados.Enabled = True
    txt_grupo_cliente.SetFocus
End Sub
Private Sub cmd_ok_Click()
    
    On Error GoTo FileError
    
    If ValidaCampos Then
        AtivaBotoes
        If lOpcao = 1 Then
            AtualTabe
            Call GravaAuditoria(1, Me.name, 10, PreparaLogTabela)
            If MovDescontoGrupoCliente.Incluir Then
                lCodigoGrupoCliente = Val(dtcboGrupoCliente.BoundText)
                lCodigoProduto = CLng(dtcboProduto.BoundText)
            Else
                MsgBox "Não foi possível incluir este registro!", vbInformation, "Erro de Verificação!"
            End If
        ElseIf lOpcao = 2 Then
            Call GravaAuditoria(1, Me.name, 10, "De:" & PreparaLogTabela)
            AtualTabe
            Call GravaAuditoria(1, Me.name, 10, "Para:" & PreparaLogTabela)
            'If MovDescontoGrupoCliente.Alterar(lCodigoGrupoCliente, lCodigoProduto) Then
            If MovDescontoGrupoCliente.Alterar(lCodigoGrupoCliente, lCodigoProduto) Then
                If GrupoCliente.Codigo > 1 Then
                'fazer a pergunta
                    Call AlteraClientesDoMesmoGrupo(GrupoCliente.Codigo, lCodigoProduto)
                End If
            Else
                MsgBox "Não foi possível alterar este registro!", vbInformation, "Erro de Verificação!"
            End If
        End If
        If MovDescontoGrupoCliente.LocalizarCodigo(CLng(dtcboGrupoCliente.BoundText), CLng(dtcboProduto.BoundText)) Then
            PesquisaPrecoPersonalizado
            AtualTela
        Else
            LimpaTela
            MsgBox "Registro não encontrado!", vbInformation, "Erro de Integridade!"
        End If
        lOpcao = 0
        cmd_novo.SetFocus
    End If
    Exit Sub
FileError:
    MsgBox Error
    Exit Sub
End Sub
Function ValidaCampos() As Boolean
    ValidaCampos = False
    If dtcboGrupoCliente.BoundText = "" Then
        MsgBox "Escolha o grupo de clientes.", vbInformation, "Atenção!"
        dtcboGrupoCliente.SetFocus
    ElseIf dtcboProduto.BoundText = "" Then
        MsgBox "Escolha o produto.", vbInformation, "Atenção!"
        dtcboProduto.SetFocus
    ElseIf fValidaValor(txt_valor_descontar.Text) = 0 And fValidaValor(txt_percentual_descontar.Text) = 0 And fValidaValor(txtPrecoFixo.Text) = 0 Then
        MsgBox "Informe o valor, percentual ou preço fixo.", vbInformation, "Atenção!"
        txt_valor_descontar.SetFocus
    ElseIf fValidaValor(txt_valor_descontar.Text) <> 0 And fValidaValor(txt_percentual_descontar.Text) <> 0 Then
        MsgBox "Apenas o valor ou percentual à descontar dever ser informado.", vbInformation, "Atenção!"
        txt_valor_descontar.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_pesquisa_Click()
'    Call GravaAuditoria(1, Me.name, 5, "")
'    ConsultaMovDescontoGrupoCliente.Show 1
'    If Len(g_string) > 0 Then
'        lCodigoGrupoCliente = RetiraGString(1)
'        lCodigoProduto = RetiraGString(2)
'        If MovDescontoGrupoCliente.LocalizarCodigo(lCodigoGrupoCliente, lCodigoProduto) Then
'            PesquisaPrecoPersonalizado
'            AtualTela
'        End If
'    End If
End Sub
Private Sub cmd_primeiro_Click()
    Call GravaAuditoria(1, Me.name, 12, "")
    If MovDescontoGrupoCliente.LocalizarPrimeiro Then
        PesquisaPrecoPersonalizado
        AtualTela
        cmd_proximo.SetFocus
    Else
        MsgBox "Não há registro nesta tabela.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    Call GravaAuditoria(1, Me.name, 14, "")
    If MovDescontoGrupoCliente.LocalizarProximo Then
        AtualTela
        PesquisaPrecoPersonalizado
    Else
        MsgBox "Fim de Arquivo.", vbInformation, "Atenção!"
        cmd_anterior.SetFocus
    End If
End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub cmd_ultimo_Click()
    Call GravaAuditoria(1, Me.name, 15, "")
    If MovDescontoGrupoCliente.LocalizarUltimo Then
        PesquisaPrecoPersonalizado
        AtualTela
        cmd_anterior.SetFocus
    Else
        MsgBox "Não há registro nesta tabela.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub dtcboGrupoCliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_produto.SetFocus
    End If
End Sub
Private Sub dtcboGrupoCliente_LostFocus()
    If dtcboGrupoCliente.BoundText <> "" And lOpcao > 0 Then
        txt_grupo_cliente.Text = dtcboGrupoCliente.BoundText
        txt_Grupo_cliente_LostFocus
        txt_produto.SetFocus
    End If
End Sub
Private Sub dtcboProduto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_percentual_descontar.SetFocus
    End If
End Sub
Private Sub dtcboProduto_LostFocus()
    If dtcboProduto.BoundText <> "" And lOpcao > 0 Then
        txt_produto.Text = dtcboProduto.BoundText
        If lOpcao = 1 Then
            If MovDescontoGrupoCliente.LocalizarCodigo(CLng(dtcboGrupoCliente.BoundText), CLng(dtcboProduto.BoundText)) Then
                MsgBox "Este produto já tem preço personalizado para este cliente." & Chr(10) & Chr(10) & "Mude o produto informado.", vbInformation, "Duplicidade de Registro!"
                txt_produto.Text = ""
                dtcboProduto.BoundText = ""
                dtcboProduto.SetFocus
                Exit Sub
            End If
        End If
        txt_produto_LostFocus
        txt_percentual_descontar.SetFocus
    End If
End Sub
Private Sub Form_Activate()
    If lFlagMovimento = 0 Then
        lOpcao = 0
        DesativaBotoes
        If MovDescontoGrupoCliente.LocalizarUltimo() Then
            AtivaBotoes
            PesquisaPrecoPersonalizado
            AtualTela
        Else
            cmd_novo.Enabled = True
            cmd_sair.Enabled = True
            PesquisaPrecoPersonalizado
        End If
        If cmd_novo.Enabled Then
            cmd_novo.SetFocus
        End If
    Else
        lFlagMovimento = 0
    End If
    Screen.MousePointer = 1
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
    Call GravaAuditoria(1, Me.name, 1, "")
    CentraForm Me
    'Set adodcGrupoCliente.Recordset = Conectar.RsConexao("SELECT Codigo, [Razao Social] FROM Cliente WHERE Inativo = " & preparaBooleano(False) & " ORDER BY [Razao Social]")
    Set adodcGrupoCliente.Recordset = Conectar.RsConexao("SELECT Codigo, Nome FROM GrupoCliente ORDER BY Nome")
    Set adodcProduto.Recordset = Conectar.RsConexao("SELECT Codigo, Nome FROM Produto WHERE Inativo = " & preparaBooleano(False) & " ORDER BY Nome")
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub grid_dados_DblClick()
    If lOpcao = 0 Then
        MarcaCelulaGrid
    End If
End Sub
Private Sub grid_dados_GotFocus()
'    grid_dados.Row = 1
'    grid_dados.Col = 0
End Sub
Private Sub grid_dados_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If lOpcao = 0 Then
            MarcaCelulaGrid
        End If
    End If
End Sub
Private Sub MarcaCelulaGrid()
    grid_dados.Col = 0
    If grid_dados.Text <> "" Then
        grid_dados.Col = 0
        lCodigoGrupoCliente = grid_dados.Text
        grid_dados.Col = 2
        lCodigoProduto = grid_dados.Text
        If MovDescontoGrupoCliente.LocalizarCodigo(lCodigoGrupoCliente, lCodigoProduto) Then
            PesquisaPrecoPersonalizado
            AtualTela
        End If
        cmd_alterar.SetFocus
    End If
End Sub
Private Sub MontaGrid()
    LimpaGrid
    grid_dados.Row = 0
    grid_dados.Col = 0
    grid_dados.Text = "Cod."
    grid_dados.ColWidth(0) = TextWidth(String$(5, "9"))
    grid_dados.ColAlignment(0) = 1
    grid_dados.Col = 1
    grid_dados.Text = "Nome"
    grid_dados.ColWidth(1) = TextWidth(String$(24, "9"))
    grid_dados.ColAlignment(1) = 0
   'obs: o "9"equivale ao tab
    '0 = left, 1 = right ,2 =  center
    grid_dados.Col = 2
    grid_dados.Text = "Cod."
    grid_dados.ColWidth(2) = TextWidth(String$(5, "9"))
    grid_dados.ColAlignment(2) = 1
    grid_dados.Col = 3
    grid_dados.Text = "Discriminação do Produto"
    grid_dados.ColWidth(3) = TextWidth(String$(25, "9"))
    grid_dados.ColAlignment(3) = 0
    grid_dados.Col = 4
    grid_dados.Text = "% à Descontar"
    grid_dados.ColWidth(4) = TextWidth(String$(12, "9"))
    grid_dados.ColAlignment(4) = 1
    grid_dados.Col = 5
    grid_dados.Text = "$ à Descontar"
    grid_dados.ColWidth(5) = TextWidth(String$(12, "9"))
    grid_dados.ColAlignment(5) = 1
    grid_dados.Col = 6
    grid_dados.Text = "$ Fixo"
    grid_dados.ColWidth(6) = TextWidth(String$(12, "9"))
    grid_dados.ColAlignment(6) = 1
    grid_dados.Col = 7
    grid_dados.Text = "Operação"
    grid_dados.ColWidth(7) = TextWidth(String$(12, "9"))
    grid_dados.ColAlignment(7) = 1
    grid_dados.Col = 8
    grid_dados.Text = "$ ECF"
    grid_dados.ColWidth(8) = TextWidth(String$(12, "9"))
    grid_dados.ColAlignment(8) = 1
End Sub
Private Sub txt_Grupo_cliente_GotFocus()
    If lOpcao = 1 Then
        txt_grupo_cliente.Text = ""
    End If
End Sub
Private Sub txt_Grupo_cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dtcboGrupoCliente.SetFocus
    End If
End Sub
Private Sub txt_Grupo_cliente_LostFocus()
    If Val(txt_grupo_cliente.Text) > 0 And lOpcao > 0 Then
        If Cliente.LocalizarCodigo(Val(txt_grupo_cliente.Text)) Then
            If Cliente.Inativo Then
                MsgBox "O cliente " & Trim(Cliente.RazaoSocial) & " está inativo.", vbInformation, "Atenção!"
                txt_grupo_cliente.SetFocus
                Exit Sub
            Else
                dtcboGrupoCliente.BoundText = Cliente.Codigo
                txt_produto.SetFocus
            End If
        Else
            MsgBox "Cliente não cadastrado.", vbInformation, "Atenção!"
            txt_grupo_cliente.SetFocus
            Exit Sub
        End If
    End If
End Sub
Private Sub txt_percentual_descontar_GotFocus()
    txt_percentual_descontar.SelStart = 0
    txt_percentual_descontar.SelLength = Len(txt_percentual_descontar.Text)
End Sub
Private Sub txt_produto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dtcboProduto.SetFocus
    End If
End Sub
Private Sub txt_produto_LostFocus()
    If fContemLetra(txt_produto) Then
        SendKeys txt_produto.Text
        txt_produto.Text = ""
    End If
    If Val(txt_produto.Text) > 0 And lOpcao > 0 Then
        If Produto.LocalizarCodigo(CLng(txt_produto.Text)) Then
            If Produto.Inativo Then
                MsgBox "O produto " & Trim(Produto.Nome) & " está inativo.", vbInformation, "Produto Inativo!"
                txt_produto.SetFocus
                Exit Sub
            Else
                dtcboProduto.BoundText = CLng(txt_produto.Text)
                txt_percentual_descontar.SetFocus
            End If
        Else
            MsgBox "Produto não cadastrado.", vbInformation, "Atenção!"
            txt_produto.SetFocus
            Exit Sub
        End If
    End If
End Sub
Private Sub txt_percentual_descontar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_valor_descontar.SetFocus
    End If
End Sub
Private Sub txt_percentual_descontar_LostFocus()
    txt_percentual_descontar.Text = Format(txt_percentual_descontar.Text, "##0.00")
End Sub
Private Sub txt_valor_descontar_GotFocus()
    txt_valor_descontar.SelStart = 0
    txt_valor_descontar.SelLength = Len(txt_valor_descontar.Text)
End Sub
Private Sub txt_valor_descontar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txtPrecoFixo.SetFocus
    End If
End Sub
Private Sub txt_valor_descontar_LostFocus()
    txt_valor_descontar.Text = Format(txt_valor_descontar.Text, "###,##0.0000")
End Sub
Private Sub txtPrecoECF_GotFocus()
    txtPrecoECF.SelStart = 0
    txtPrecoECF.SelLength = Len(txtPrecoECF.Text)
End Sub
Private Sub txtPrecoECF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub txtPrecoECF_LostFocus()
    txtPrecoECF.Text = Format(txtPrecoECF.Text, "###,##0.0000")
End Sub
Private Sub txtPrecoFixo_GotFocus()
    txtPrecoFixo.SelStart = 0
    txtPrecoFixo.SelLength = Len(txtPrecoFixo.Text)
End Sub
Private Sub txtPrecoFixo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub txtPrecoFixo_LostFocus()
    txtPrecoFixo.Text = Format(txtPrecoFixo.Text, "###,##0.0000")
End Sub
