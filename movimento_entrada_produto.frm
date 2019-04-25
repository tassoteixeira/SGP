VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form movimento_entrada_produto 
   Caption         =   "Movimento de Entrada de Produtos"
   ClientHeight    =   6990
   ClientLeft      =   165
   ClientTop       =   285
   ClientWidth     =   8550
   Icon            =   "movimento_entrada_produto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   Picture         =   "movimento_entrada_produto.frx":030A
   ScaleHeight     =   6990
   ScaleWidth      =   8550
   Begin VB.Frame frmDados 
      Height          =   3435
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8295
      Begin MSAdodcLib.Adodc adodcProduto 
         Height          =   330
         Left            =   2400
         Top             =   1680
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
      Begin MSAdodcLib.Adodc adodc_fornecedor 
         Height          =   330
         Left            =   1680
         Top             =   1080
         Visible         =   0   'False
         Width           =   2955
         _ExtentX        =   5212
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
         Caption         =   "adodc_fornecedor"
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
      Begin VB.TextBox txt_preco_custo 
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   15
         Top             =   2310
         Width           =   1095
      End
      Begin VB.TextBox txt_quantidade 
         Height          =   285
         Left            =   3000
         MaxLength       =   10
         TabIndex        =   17
         Top             =   2310
         Width           =   1095
      End
      Begin VB.TextBox txt_observacao 
         Height          =   285
         Left            =   3000
         MaxLength       =   40
         TabIndex        =   23
         Top             =   2970
         Width           =   5175
      End
      Begin VB.TextBox txt_preco_total 
         Height          =   285
         Left            =   6000
         MaxLength       =   10
         TabIndex        =   19
         Top             =   2310
         Width           =   1095
      End
      Begin VB.TextBox txt_produto 
         Height          =   285
         Left            =   120
         MaxLength       =   18
         TabIndex        =   10
         Top             =   1680
         Width           =   795
      End
      Begin VB.ComboBox cbo_tipo_entrada 
         Height          =   315
         Left            =   6000
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   420
         Width           =   2175
      End
      Begin VB.TextBox txt_numero_documento 
         Height          =   285
         Left            =   3000
         MaxLength       =   10
         TabIndex        =   4
         Top             =   420
         Width           =   1095
      End
      Begin MSMask.MaskEdBox msk_data_entrada 
         Height          =   300
         Left            =   120
         TabIndex        =   2
         Top             =   420
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSDataListLib.DataCombo dtcbo_fornecedor 
         Bindings        =   "movimento_entrada_produto.frx":0750
         DataSource      =   "adodc_fornecedor"
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   "dtcbo_fornecedor"
      End
      Begin MSDataListLib.DataCombo dtcboProduto 
         Bindings        =   "movimento_entrada_produto.frx":076F
         Height          =   315
         Left            =   960
         TabIndex        =   11
         Top             =   1680
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   "dtcboProduto"
      End
      Begin MSMask.MaskEdBox msk_data_digitacao 
         Height          =   300
         Left            =   120
         TabIndex        =   21
         Top             =   2970
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         Caption         =   "Total da &Entrada"
         Height          =   195
         Index           =   5
         Left            =   6000
         TabIndex        =   18
         Top             =   2100
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "&Preço de Custo"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   2100
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "&Quantidade"
         Height          =   195
         Index           =   2
         Left            =   3000
         TabIndex        =   16
         Top             =   2100
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "D&ata da digitação"
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   20
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "O&bservação"
         Height          =   195
         Index           =   9
         Left            =   3000
         TabIndex        =   22
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "P&roduto"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   1470
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Unidade"
         Height          =   195
         Index           =   3
         Left            =   6000
         TabIndex        =   12
         Top             =   1470
         Width           =   675
      End
      Begin VB.Label lbl_unidade 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   6000
         TabIndex        =   13
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "&Fornecedor"
         Height          =   300
         Index           =   10
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "&Data da entrada"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   1
         Top             =   180
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "&Tipo da entrada"
         Height          =   195
         Index           =   7
         Left            =   6000
         TabIndex        =   5
         Top             =   210
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "&Numero do documento"
         Height          =   195
         Index           =   6
         Left            =   3000
         TabIndex        =   3
         Top             =   210
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   120
      Picture         =   "movimento_entrada_produto.frx":078A
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Cria um novo registro."
      Top             =   6060
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   1020
      Picture         =   "movimento_entrada_produto.frx":1E1C
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Altera o registro atual."
      Top             =   6060
      Width           =   795
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   1920
      Picture         =   "movimento_entrada_produto.frx":3316
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Exclui o registro atual."
      Top             =   6060
      Width           =   795
   End
   Begin VB.CommandButton cmd_pesquisa 
      Caption         =   "&Pesquisa"
      Height          =   855
      Left            =   2820
      Picture         =   "movimento_entrada_produto.frx":49A8
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Pesquisa um registro específico."
      Top             =   6060
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3720
      Picture         =   "movimento_entrada_produto.frx":5E1A
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   6060
      Width           =   795
   End
   Begin MSGrid.Grid grid_oleo 
      Height          =   2415
      Left            =   120
      TabIndex        =   36
      Top             =   3540
      Width           =   8295
      _Version        =   65536
      _ExtentX        =   14631
      _ExtentY        =   4260
      _StockProps     =   77
      BackColor       =   16777215
      Cols            =   9
      FixedCols       =   0
   End
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   6240
      TabIndex        =   31
      Top             =   5940
      Width           =   2175
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "movimento_entrada_produto.frx":74AC
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "movimento_entrada_produto.frx":89A6
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "movimento_entrada_produto.frx":9EA0
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "movimento_entrada_produto.frx":B312
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   7620
      Picture         =   "movimento_entrada_produto.frx":C894
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Cancela o registro atual."
      Top             =   6060
      Width           =   795
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   6720
      Picture         =   "movimento_entrada_produto.frx":DD8E
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Confirma o registro atual."
      Top             =   6060
      Width           =   795
   End
End
Attribute VB_Name = "movimento_entrada_produto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flag_movimento_entrada_produto As Integer
Dim lOpcao As String
Dim lEmpresa As Integer
Dim lData As Date
Dim lCodigoProduto As Long
Dim lNumeroDocumento As String
Dim lQuantidade As Currency
Dim lDataDigitacao As Date
Dim lSQL As String
Dim lGravados As Long

Dim rsDadosGrid As New adodb.Recordset
Dim fld As adodb.Field

Private EntradaProduto As New cEntradaProduto
Private EntradaProdutoCabecalho As New cEntradaProdutoCabecalho
Private Estoque As New cEstoque
Private Produto As New cProduto
Private SubEstoque As New cSubEstoque

Private Sub PreencheGrid()
    Dim i As Integer
    Dim i2 As Integer
    Dim xLinha As Integer
    xLinha = 0
    If Not rsDadosGrid.EOF Then
        rsDadosGrid.MoveFirst
        Do Until rsDadosGrid.EOF
            xLinha = xLinha + 1
            If xLinha > 1 Then
                grid_oleo.Rows = grid_oleo.Rows + 1
            End If
            grid_oleo.Row = xLinha
            i2 = -1
            For Each fld In rsDadosGrid.Fields
                i2 = i2 + 1
                grid_oleo.Col = i2
                'If fld.Name = "CODIGO" Then
                '    Grid1.Text = fMascaraContaContabil(fld.Value)
                'Else
                    If IsNull(fld.Value) Then
                        grid_oleo.Text = ""
                    Else
                        If fld.Type = adCurrency Then
                            grid_oleo.Text = Format(fld.Value, "###,###,##0.00")
                        Else
                            grid_oleo.Text = fld.Value
                        End If
                    End If
                'End If
            Next
            rsDadosGrid.MoveNext
        Loop
        grid_oleo.Row = grid_oleo.Rows - 1
        grid_oleo.Col = 0
    End If
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
    If lOpcao = 0 Then
        VerificaLiberacaoDigitacao
    End If
End Sub
Private Sub Inclui()
    lOpcao = 1
    DesativaBotoes
    cmd_novo.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
End Sub
Private Sub AtualizaGrid()
    If Val(lNumeroDocumento) = 0 Then
        lNumeroDocumento = 0
        lData = g_data_def
    End If
    lSQL = "Select Entrada_Produto.[Data da Entrada], Entrada_Produto.[Numero do Documento], Produto.Nome, Entrada_Produto.[Preco de Custo], Entrada_Produto.Quantidade, Entrada_Produto.[Tipo da Entrada], Entrada_Produto.[Data da Digitacao], Entrada_Produto.Observacao, Entrada_Produto.[Codigo do Produto]"
    lSQL = lSQL & " From Entrada_Produto, Produto"
    lSQL = lSQL & " Where Entrada_Produto.Empresa = " & g_empresa
    lSQL = lSQL & " And Entrada_Produto.[Data da Entrada] = " & preparaData(lData)
    lSQL = lSQL & " And Entrada_Produto.[Numero do Documento] = " & preparaTexto(lNumeroDocumento)
    lSQL = lSQL & " And Produto.Codigo = Entrada_Produto.[Codigo do Produto]"
    lSQL = lSQL & " Order by [Data da Entrada], [Codigo do Produto], [Numero do Documento]"
    Set rsDadosGrid = Conectar.RsConexao(lSQL)
    FormataGrid
    PreencheGrid
'    If rsDadosGrid.RecordCount > 0 Then
'        rsDadosGrid.MoveFirst
'        Do Until rsDadosGrid.EOF
'            AdcionaDadosGrid
'            rsDadosGrid.MoveNext
'        Loop
'    End If
End Sub
Private Sub AtualTabe()
    EntradaProduto.Empresa = g_empresa
    EntradaProduto.DataEntrada = Format(msk_data_entrada.Text, "dd/mm/yyyy")
    EntradaProduto.CodigoProduto = CLng(dtcboProduto.BoundText)
    EntradaProduto.NumeroDocumento = txt_numero_documento.Text
    EntradaProduto.TipoEntrada = cbo_tipo_entrada.ItemData(cbo_tipo_entrada.ListIndex)
    EntradaProduto.PrecoCusto = fValidaValor(txt_preco_custo.Text)
    EntradaProduto.Quantidade = fValidaValor(txt_quantidade.Text)
    EntradaProduto.TotalCusto = fValidaValor(txt_preco_total.Text)
    EntradaProduto.CodigoFornecedor = Val(dtcbo_fornecedor.BoundText)
    EntradaProduto.DataDigitacao = Format(msk_data_digitacao.Text, "dd/mm/yyyy")
    EntradaProduto.Observacao = txt_observacao.Text
    If Produto.PrecoCusto <> EntradaProduto.PrecoCusto Then
        If MsgBox("O preço de custo informado nesta nota é: " & Format(EntradaProduto.PrecoCusto, "###,###,##0.0000") & Chr(10) & Chr(13) & Chr(10) & "E o preço de custo no cadastro do produto é: " & Format(Produto.PrecoCusto, "###,###,##0.0000") & Chr(10) & Chr(13) & Chr(10) & "Deseja alterar o preço de custo no cadastro de produtos?", vbYesNo + vbDefaultButton2, "Alteração de Preço de Custo") = vbYes Then
            Produto.PrecoCusto = EntradaProduto.PrecoCusto
            If Not Produto.AlterarCusto(EntradaProduto.CodigoProduto, g_empresa, Produto.PrecoCusto) Then
                MsgBox "Não foi possível alterar o produto!", vbInformation, "Erro de Integridade"
            End If
        End If
    End If
    EntradaProduto.CodigoUsuario = g_usuario
    EntradaProduto.CustoUnitarioBruto = fValidaValor(txt_preco_custo.Text)
    EntradaProduto.SubEstoque = 1

    EntradaProdutoCabecalho.Empresa = g_empresa
    EntradaProdutoCabecalho.DataEntrada = Format(msk_data_entrada.Text, "dd/mm/yyyy")
    EntradaProdutoCabecalho.NumeroDocumento = txt_numero_documento.Text
    EntradaProdutoCabecalho.CodigoFornecedor = Val(dtcbo_fornecedor.BoundText)
    EntradaProdutoCabecalho.TipoEntrada = cbo_tipo_entrada.ItemData(cbo_tipo_entrada.ListIndex)
    EntradaProdutoCabecalho.SubEstoque = 1
    EntradaProdutoCabecalho.TotalProduto = fValidaValor(txt_preco_total.Text)
    EntradaProdutoCabecalho.Desconto = 0
    EntradaProdutoCabecalho.Substituicao = 0
    EntradaProdutoCabecalho.Outros = 0
    EntradaProdutoCabecalho.TotalNota = fValidaValor(txt_preco_total.Text)
    EntradaProdutoCabecalho.DataDigitacao = Format(msk_data_digitacao.Text, "dd/mm/yyyy")
    EntradaProdutoCabecalho.CodigoUsuario = g_usuario
    EntradaProdutoCabecalho.Observacao = txt_observacao.Text
End Sub
Private Sub AtualTela()
    Dim i As Integer
    
    lData = EntradaProduto.DataEntrada
    lCodigoProduto = EntradaProduto.CodigoProduto
    lNumeroDocumento = EntradaProduto.NumeroDocumento
    lQuantidade = EntradaProduto.Quantidade
    lDataDigitacao = EntradaProduto.DataDigitacao
    
    msk_data_entrada.Text = Format(EntradaProduto.DataEntrada, "dd/mm/yyyy")
    txt_numero_documento.Text = EntradaProduto.NumeroDocumento
    cbo_tipo_entrada.ListIndex = EntradaProduto.TipoEntrada - 1
    dtcbo_fornecedor.BoundText = EntradaProduto.CodigoFornecedor
    dtcboProduto.BoundText = ""
    lbl_unidade.Caption = ""
    If Produto.LocalizarCodigo(EntradaProduto.CodigoProduto) Then
        txt_produto.Text = EntradaProduto.CodigoProduto
        dtcboProduto.BoundText = EntradaProduto.CodigoProduto
        lbl_unidade.Caption = Produto.Unidade
    End If
    txt_preco_custo.Text = Format(EntradaProduto.PrecoCusto, "###,##0.0000")
    txt_quantidade.Text = Format(EntradaProduto.Quantidade, "###,##0.00;-###,##0.00")
    txt_preco_total.Text = Format(EntradaProduto.TotalCusto, "###,##0.00")
    msk_data_digitacao.Text = Format(EntradaProduto.DataDigitacao, "dd/mm/yyyy")
    txt_observacao.Text = EntradaProduto.Observacao
    frmDados.Enabled = False
    VerificaLiberacaoDigitacao
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
    Set EntradaProduto = Nothing
    Set EntradaProdutoCabecalho = Nothing
    Set Estoque = Nothing
    Set Produto = Nothing
    Set SubEstoque = Nothing
End Sub
Private Sub PreencheCboTipoEntrada()
    cbo_tipo_entrada.Clear
    cbo_tipo_entrada.AddItem "1 Normal"
    cbo_tipo_entrada.ItemData(cbo_tipo_entrada.NewIndex) = 1
    cbo_tipo_entrada.AddItem "2 Acerto de Estoque"
    cbo_tipo_entrada.ItemData(cbo_tipo_entrada.NewIndex) = 2
    cbo_tipo_entrada.AddItem "3 Inventário"
    cbo_tipo_entrada.ItemData(cbo_tipo_entrada.NewIndex) = 3
    cbo_tipo_entrada.AddItem "4 Transferência"
    cbo_tipo_entrada.ItemData(cbo_tipo_entrada.NewIndex) = 4
End Sub
Private Sub cbo_tipo_entrada_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dtcbo_fornecedor.SetFocus
    End If
End Sub
Private Sub cmd_alterar_Click()
'    zzMoveEntradaParaEstoque
'    zzMudaCodigoProduto
'    Exit Sub
'
    Call GravaAuditoria(1, Me.name, 3, "")
    lOpcao = 2
    DesativaBotoes
    cmd_alterar.Visible = True
    cmd_alterar.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    frmDados.Enabled = True
    txt_quantidade.SetFocus
End Sub
Private Sub cmd_anterior_Click()
    Call GravaAuditoria(1, Me.name, 13, "")
    If EntradaProduto.LocalizarAnterior Then
        AtualTela
    Else
        MsgBox "Início de Arquivo.", 48, "Atenção!"
        cmd_proximo.SetFocus
    End If
End Sub
Private Sub cmd_cancelar_Click()
    Call GravaAuditoria(1, Me.name, 9, "")
    LimpaTela
    If EntradaProduto.LocalizarCodigo(g_empresa, lData, lCodigoProduto, lNumeroDocumento) Then
        AtualTela
        AtivaBotoes
        AtualizaGrid
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
    Dim i As Integer
    Do Until grid_oleo.Rows = 2
        grid_oleo.Row = grid_oleo.Rows - 1
        grid_oleo.RemoveItem grid_oleo.Row
    Loop
    grid_oleo.Row = 1
    For i = 0 To 8
        grid_oleo.Col = i
        grid_oleo.Text = ""
    Next
End Sub
Private Sub LimpaTela()
    If lGravados = 0 Then
        msk_data_entrada.Text = "__/__/____"
        txt_numero_documento.Text = ""
        cbo_tipo_entrada.ListIndex = -1
        dtcbo_fornecedor.BoundText = ""
    End If
    txt_produto.Text = ""
    dtcboProduto.BoundText = ""
    lbl_unidade.Caption = ""
    txt_preco_custo.Text = ""
    txt_quantidade.Text = ""
    txt_preco_total.Text = ""
    msk_data_digitacao.Text = "__/__/____"
    txt_observacao.Text = ""
End Sub
Private Sub cmd_excluir_Click()
    Call GravaAuditoria(1, Me.name, 4, "")
    If Val(txt_produto.Text) > 0 Then
        If (MsgBox("Deseja excluir este registro?", 4 + 32 + 256, "Exclusão de Registro!")) = 6 Then
            Call GravaAuditoria(1, Me.name, 10, "Dt:" & msk_data_entrada.Text & " N.Doc:" & txt_numero_documento.Text & " Tp:" & cbo_tipo_entrada.Text & " Prod:" & txt_produto.Text & " Qtd:" & txt_quantidade.Text & " Vlr.Unit:" & txt_preco_custo.Text)
            lOpcao = 3
            Conectar.IniciaTransacao
            If EntradaProduto.Excluir(g_empresa, CDate(msk_data_entrada.Text), CLng(txt_produto.Text), txt_numero_documento.Text) Then
                If Estoque.AlterarQuantidade(g_empresa, lCodigoProduto, lQuantidade, False) Then
                    If SubEstoque.AlterarQuantidade(g_empresa, lCodigoProduto, 1, lQuantidade, False) Then
                        Conectar.ConfirmaTransacao
                    Else
                        Conectar.CancelaTransacao
                        MsgBox "Não foi possível alterar o Sub-Estoque!", vbInformation, "Erro de Integridade!"
                    End If
                Else
                    Conectar.CancelaTransacao
                    MsgBox "Não foi possível alterar o Estoque!", vbInformation, "Erro de Integridade!"
                End If
            Else
                Conectar.CancelaTransacao
                MsgBox "Não foi possível excluir este registro!", vbInformation, "Erro de Integridade!"
            End If
            TotalizaNotaCabecalho
            If EntradaProdutoCabecalho.TotalNota = 0 Then
                If Not EntradaProdutoCabecalho.Excluir(g_empresa, lData, lNumeroDocumento) Then
                    MsgBox "Não foi possível excluir o registro de cabecalho!", vbInformation, "Erro de Integridade!"
                End If
            End If
            If EntradaProduto.LocalizarUltimo(g_empresa) Then
                AtualTela
                AtivaBotoes
            Else
                LimpaTela
                DesativaBotoes
                cmd_novo.Enabled = True
                cmd_sair.Enabled = True
                cmd_novo.SetFocus
            End If
            lOpcao = 0
            AtualizaGrid
        End If
    End If
End Sub
Private Sub cmd_novo_Click()
'    zzAlteraPrecoCusto
'    Exit Sub
'
    Call GravaAuditoria(1, Me.name, 2, "")
    LimpaTela
    Inclui
    frmDados.Enabled = True
    If lGravados = 0 Then
        msk_data_entrada.SetFocus
    Else
        dtcboProduto.SetFocus
        If cbo_tipo_entrada.ItemData(cbo_tipo_entrada.ListIndex) = 3 Then
            txt_produto.SetFocus
        End If
    End If
End Sub
Private Sub cmd_ok_Click()
    Dim xAtualizou As Boolean
    On Error GoTo FileError
    
    If ValidaCampos Then
        If VerificaLiberacaoDigitacao2 Then
            AtivaBotoes
            If lOpcao = 1 Then
                Call GravaAuditoria(1, Me.name, 10, "Dt:" & msk_data_entrada.Text & " N.Doc:" & txt_numero_documento.Text & " Tp:" & cbo_tipo_entrada.Text & " Prod:" & txt_produto.Text & " Qtd:" & txt_quantidade.Text & " Vlr.Unit:" & txt_preco_custo.Text)
                AtualTabe
                Conectar.IniciaTransacao
                If EntradaProduto.Incluir Then
                    lData = CDate(msk_data_entrada.Text)
                    lCodigoProduto = CLng(dtcboProduto.BoundText)
                    lNumeroDocumento = txt_numero_documento.Text
                    lQuantidade = fValidaValor(txt_quantidade.Text)
                    lGravados = lGravados + 1
                    If lGravados = 1 Then
                        If Not EntradaProdutoCabecalho.Incluir Then
                            MsgBox "Não foi possível incluir o registro de cabecalho!", vbInformation, "Erro de Integridade!"
                        End If
                    End If
                    If Estoque.AlterarQuantidade(g_empresa, lCodigoProduto, lQuantidade, True) Then
                        If SubEstoque.AlterarQuantidade(g_empresa, lCodigoProduto, 1, lQuantidade, True) Then
                            Conectar.ConfirmaTransacao
                        Else
                            Conectar.CancelaTransacao
                            MsgBox "Não foi possível alterar o Sub-Estoque!", vbInformation, "Erro de Integridade!"
                        End If
                    Else
                        Conectar.CancelaTransacao
                        MsgBox "Não foi possível alterar o Estoque!", vbInformation, "Erro de Integridade!"
                    End If
                Else
                    Conectar.CancelaTransacao
                    MsgBox "Não foi possível incluir este registro!", vbInformation, "Erro de Integridade!"
                End If
            ElseIf lOpcao = 2 Then
                xAtualizou = False
                Call GravaAuditoria(1, Me.name, 10, "De: Dt:" & Format(EntradaProduto.DataEntrada, "dd/mm/yyyy") & " N.Doc:" & EntradaProduto.NumeroDocumento & " Tp:" & EntradaProduto.TipoEntrada & " Prod:" & EntradaProduto.CodigoProduto & " Qtd:" & EntradaProduto.Quantidade & " Vlr.Unit:" & EntradaProduto.PrecoCusto)
                Call GravaAuditoria(1, Me.name, 10, "Para: Dt:" & msk_data_entrada.Text & " N.Doc:" & txt_numero_documento.Text & " Tp:" & cbo_tipo_entrada.Text & " Prod:" & txt_produto.Text & " Qtd:" & txt_quantidade.Text & " Vlr.Unit:" & txt_preco_custo.Text)
                Conectar.IniciaTransacao
                If Estoque.AlterarQuantidade(g_empresa, lCodigoProduto, lQuantidade, False) Then
                    If SubEstoque.AlterarQuantidade(g_empresa, lCodigoProduto, 1, lQuantidade, False) Then
                        xAtualizou = True
                    Else
                        Conectar.CancelaTransacao
                        MsgBox "Não foi possível alterar o Sub-Estoque!", vbInformation, "Erro de Integridade!"
                    End If
                Else
                    Conectar.CancelaTransacao
                    MsgBox "Não foi possível alterar o Estoque!", vbInformation, "Erro de Integridade!"
                End If
                If xAtualizou Then
                    xAtualizou = False
                    AtualTabe
                    If EntradaProduto.Alterar(g_empresa, lData, lCodigoProduto, lNumeroDocumento) Then
                        If EntradaProdutoCabecalho.Alterar(g_empresa, lData, lNumeroDocumento) Then
                            lData = CDate(msk_data_entrada.Text)
                            lCodigoProduto = CLng(dtcboProduto.BoundText)
                            lNumeroDocumento = txt_numero_documento.Text
                        Else
                            MsgBox "Não foi possível alterar o registro de cabecalho!", vbInformation, "Erro de Integridade!"
                        End If
                        lData = CDate(msk_data_entrada.Text)
                        lCodigoProduto = CLng(dtcboProduto.BoundText)
                        lNumeroDocumento = txt_numero_documento.Text
                        lQuantidade = fValidaValor(txt_quantidade.Text)
                        If Estoque.AlterarQuantidade(g_empresa, lCodigoProduto, lQuantidade, True) Then
                            If SubEstoque.AlterarQuantidade(g_empresa, lCodigoProduto, 1, lQuantidade, True) Then
                                Conectar.ConfirmaTransacao
                            Else
                                Conectar.CancelaTransacao
                                MsgBox "Não foi possível alterar o Sub-Estoque!", vbInformation, "Erro de Integridade!"
                            End If
                        Else
                            Conectar.CancelaTransacao
                            MsgBox "Não foi possível alterar o Estoque!", vbInformation, "Erro de Integridade!"
                        End If
                    Else
                        Conectar.CancelaTransacao
                        MsgBox "Não foi possível alterar este registro!", vbInformation, "Erro de Integridade!"
                    End If
                End If
            End If
            TotalizaNotaCabecalho
            If EntradaProduto.LocalizarCodigo(g_empresa, lData, lCodigoProduto, lNumeroDocumento) Then
                AtualTela
                AtualizaGrid
            End If
            cmd_novo.SetFocus
            If lOpcao = 1 Then
                lOpcao = 0
                cmd_novo_Click
            Else
                lOpcao = 0
            End If
        End If
    End If
    Exit Sub
FileError:
    Exit Sub
End Sub
Private Sub TotalizaNotaCabecalho()
    Dim rsTotalNota As New adodb.Recordset
    Dim xTotal As Currency
    
    xTotal = 0
    If Val(lNumeroDocumento) = 0 Then
        lNumeroDocumento = 0
        lData = g_data_def
    End If
    lSQL = "SELECT [Total do Custo]"
    lSQL = lSQL & " FROM  Entrada_Produto"
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    lSQL = lSQL & " AND   [Data da Entrada] = " & preparaData(lData)
    lSQL = lSQL & " AND   [Numero do Documento] = " & preparaTexto(lNumeroDocumento)
    Set rsTotalNota = Conectar.RsConexao(lSQL)
    If rsTotalNota.RecordCount > 0 Then
        rsTotalNota.MoveFirst
        Do Until rsTotalNota.EOF
            xTotal = xTotal + rsTotalNota("Total do Custo").Value
            rsTotalNota.MoveNext
        Loop
    End If
    If EntradaProdutoCabecalho.LocalizarCodigo(g_empresa, lData, lNumeroDocumento) Then
        EntradaProdutoCabecalho.TotalNota = xTotal
        EntradaProdutoCabecalho.TotalProduto = xTotal
        If Not EntradaProdutoCabecalho.Alterar(g_empresa, lData, lNumeroDocumento) Then
            MsgBox "Não foi possível alterar Cabecalho de Nota!", vbCritical, "Erro de Integridade!"
        End If
    Else
        MsgBox "Cabecalho de Nota não encontrado!", vbCritical, "Erro de Integridade!"
    End If
End Sub
Function ValidaCampos() As Boolean
    ValidaCampos = False
    If Not IsDate(msk_data_entrada.Text) Then
        MsgBox "Informe a data da entrada.", vbInformation, "Atenção!"
        msk_data_entrada.SetFocus
    ElseIf Trim(txt_numero_documento.Text) = "" Then
        MsgBox "Informe o número do documento.", vbInformation, "Atenção!"
        txt_numero_documento.SetFocus
    ElseIf cbo_tipo_entrada.ListIndex = -1 Then
        MsgBox "Escolha o tipo da entrada.", vbInformation, "Atenção!"
        cbo_tipo_entrada.SetFocus
    ElseIf Not Val(dtcbo_fornecedor.BoundText) > 0 Then
        MsgBox "Selecione o fornecedor.", vbInformation, "Atenção!"
        dtcbo_fornecedor.SetFocus
    ElseIf dtcboProduto.BoundText = "" Then
        MsgBox "Escolha o produto.", vbInformation, "Atenção!"
        dtcboProduto.SetFocus
    ElseIf Not fValidaValor(txt_preco_custo.Text) > 0 Then
        MsgBox "Informe o preço de custo.", vbInformation, "Atenção!"
        txt_preco_custo.SetFocus
    ElseIf Not fValidaValor(txt_quantidade.Text) > 0 And cbo_tipo_entrada.ListIndex <> 1 And cbo_tipo_entrada.ListIndex <> 2 Then
        MsgBox "Informe a quantidade.", vbInformation, "Atenção!"
        txt_quantidade.SetFocus
    ElseIf Not fValidaValor(txt_preco_total.Text) > 0 And cbo_tipo_entrada.ListIndex <> 1 And cbo_tipo_entrada.ListIndex <> 2 Then
        MsgBox "Informe o preço total de custo.", vbInformation, "Atenção!"
        txt_preco_total.SetFocus
    ElseIf Not IsDate(msk_data_digitacao.Text) Then
        MsgBox "Informe a data da digitação.", vbInformation, "Atenção!"
        msk_data_digitacao.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub VerificaLiberacaoDigitacao()
    Dim x_flag As Boolean
    x_flag = True
    If g_nivel_acesso > 4 Then
        If EntradaProduto.Empresa < g_cfg_empresa_i Or EntradaProduto.Empresa > g_cfg_empresa_f Then
            x_flag = False
'        ElseIf EntradaProduto.DataEntrada < g_cfg_data_i Or EntradaProduto.DataEntrada > g_cfg_data_f Then
'            x_flag = False
        End If
    End If
    If EntradaProduto.TipoEntrada = 4 Then
        x_flag = False
    End If
    If x_flag Then
        cmd_alterar.Enabled = True
        cmd_excluir.Enabled = True
    Else
        cmd_alterar.Enabled = False
        cmd_excluir.Enabled = False
    End If
End Sub
Function VerificaLiberacaoDigitacao2() As Boolean
    VerificaLiberacaoDigitacao2 = False
    If g_nivel_acesso <= 4 Then
        VerificaLiberacaoDigitacao2 = True
        Exit Function
    End If
'    If msk_data_entrada < g_cfg_data_i Or msk_data_entrada > g_cfg_data_f Then
'        MsgBox "A data da entrada deve estar entre " & Format(g_cfg_data_i, "dd/mm/yyyy") & " a " & Format(g_cfg_data_f, "dd/mm/yyyy") & ".", vbInformation, "Digitação Não Autorizada!"
'        msk_data_entrada.SetFocus
    If cbo_tipo_entrada.ListIndex = 4 Then
        MsgBox "O tipo de entrada não pode ser 4.", vbInformation, "Digitação Não Autorizada!"
        cbo_tipo_entrada.SetFocus
    Else
        VerificaLiberacaoDigitacao2 = True
    End If
End Function
Private Sub zzAlteraPrecoCusto()
    Dim rsMovimento As New adodb.Recordset
    Dim xSQL As String
    
    Exit Sub
    xSQL = "SELECT [Codigo do Produto], [Preco de Custo]"
    xSQL = xSQL & "  FROM Entrada_Produto"
    xSQL = xSQL & " WHERE Empresa = " & g_empresa
    xSQL = xSQL & "   AND [Data da Entrada] = " & preparaData(CDate("08/07/2008"))
    Set rsMovimento = Conectar.RsConexao(xSQL)
    If rsMovimento.RecordCount > 0 Then
        Do Until rsMovimento.EOF
            If Produto.LocalizarCodigo(rsMovimento("Codigo do Produto").Value) Then
                Produto.PrecoCusto = rsMovimento("Preco de Custo").Value
                If Not Produto.AlterarCusto(rsMovimento("Codigo do Produto").Value, g_empresa, Produto.PrecoCusto) Then
                    MsgBox "Nao foi possível alterar Produto"
                End If
            Else
                MsgBox "Produto nao encontrado"
            End If
            rsMovimento.MoveNext
        Loop
    End If
    rsMovimento.Close
    Set rsMovimento = Nothing
End Sub
Private Sub cmd_pesquisa_Click()
    Call GravaAuditoria(1, Me.name, 5, "")
    consulta_movimento_entrada_produto.Show 1
    If Len(g_string) > 0 Then
        lData = RetiraGString(1)
        lCodigoProduto = RetiraGString(2)
        lNumeroDocumento = RetiraGString(3)
        If EntradaProduto.LocalizarCodigo(g_empresa, lData, lCodigoProduto, lNumeroDocumento) Then
            AtualTela
            AtualizaGrid
        End If
    End If
End Sub
Private Sub cmd_primeiro_Click()
    Call GravaAuditoria(1, Me.name, 12, "")
    If EntradaProduto.LocalizarPrimeiro Then
        AtualTela
        cmd_proximo.SetFocus
    Else
        MsgBox "Não há registro nesta empresa.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    Call GravaAuditoria(1, Me.name, 14, "")
    If EntradaProduto.LocalizarProximo Then
        AtualTela
    Else
        MsgBox "Fim de Arquivo.", 48, "Atenção!"
        cmd_anterior.SetFocus
    End If
End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub cmd_ultimo_Click()
    Call GravaAuditoria(1, Me.name, 15, "")
    If EntradaProduto.LocalizarUltimo(g_empresa) Then
        AtualTela
        cmd_anterior.SetFocus
    Else
        MsgBox "Não há registro nesta empresa.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub dtcbo_fornecedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txt_produto.SetFocus
    End If
End Sub
Private Sub dtcboProduto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txt_quantidade.SetFocus
    End If
End Sub
Private Sub dtcboProduto_LostFocus()
    If dtcboProduto.BoundText <> "" And lOpcao > 0 Then
        txt_produto.Text = dtcboProduto.BoundText
        If lOpcao = 1 Then
            If ExisteRegistro Then
                txt_produto.Text = ""
                dtcboProduto.BoundText = ""
                txt_produto.SetFocus
                Exit Sub
            End If
        End If
        txt_produto_LostFocus
        txt_preco_custo.SetFocus
        If cbo_tipo_entrada.ItemData(cbo_tipo_entrada.ListIndex) = 3 Then
            txt_quantidade.SetFocus
        End If
    End If
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    If g_empresa <> lEmpresa Then
        flag_movimento_entrada_produto = 0
    End If
    If flag_movimento_entrada_produto = 0 Then
        lGravados = 0
        lOpcao = 0
        lEmpresa = g_empresa
        DesativaBotoes
        If EntradaProduto.LocalizarUltimo(g_empresa) Then
            AtualTela
            AtivaBotoes
            AtualizaGrid
        Else
            cmd_novo.Enabled = True
            cmd_sair.Enabled = True
            AtualizaGrid
        End If
        cmd_novo.SetFocus
    Else
        flag_movimento_entrada_produto = 0
    End If
    Screen.MousePointer = 1
End Sub
Private Sub Form_Deactivate()
    flag_movimento_entrada_produto = 1
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
    CentraForm Me
    
    PreencheCboTipoEntrada
    Set adodc_fornecedor.Recordset = Conectar.RsConexao("SELECT Codigo, Nome FROM Fornecedor WHERE Empresa = " & g_empresa & " ORDER BY Nome")
    Set adodcProduto.Recordset = Conectar.RsConexao("SELECT Codigo, Nome FROM Produto WHERE Inativo = " & preparaBooleano(False) & " ORDER BY Nome")
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub grid_oleo_DblClick()
    MarcaCelulas
End Sub
Private Sub grid_oleo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        MarcaCelulas
    End If
End Sub
Private Sub msk_data_digitacao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_observacao.SetFocus
    End If
End Sub
Private Sub msk_data_digitacao_LostFocus()
    If IsDate(msk_data_digitacao.Text) Then
        lDataDigitacao = msk_data_digitacao.Text
    End If
End Sub
Private Sub msk_data_entrada_GotFocus()
    If Not IsDate(msk_data_entrada.Text) Then
        msk_data_entrada = Format(g_data_def, "dd/mm/yyyy")
        lDataDigitacao = g_data_def
        txt_numero_documento = lNumeroDocumento
        cbo_tipo_entrada.ListIndex = 0
    End If
    msk_data_entrada.SelStart = 0
    msk_data_entrada.SelLength = 5
End Sub
Private Sub msk_data_entrada_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_numero_documento.SetFocus
    End If
End Sub
Private Sub MarcaCelulas()
    grid_oleo.Col = 0
    If grid_oleo.Text <> "" Then
        lData = CDate(grid_oleo.Text)
        grid_oleo.Col = 8
        lCodigoProduto = Val(grid_oleo.Text)
        grid_oleo.Col = 1
        lNumeroDocumento = grid_oleo.Text
        If EntradaProduto.LocalizarCodigo(g_empresa, lData, lCodigoProduto, lNumeroDocumento) Then
            AtualTela
        End If
        cmd_alterar.SetFocus
    End If
End Sub
Private Sub FormataGrid()
    Dim i As Integer
    
    LimpaGrid
    grid_oleo.Row = 0
    i = 0
    grid_oleo.Col = i
    grid_oleo.Text = "Data da Entrada"
    grid_oleo.ColWidth(i) = 1000 'TextWidth(String$(11, "9"))
    grid_oleo.ColAlignment(i) = 2
   'obs: o "9"equivale ao tab
    '0=left, 1=right, 2=center
    i = i + 1
    grid_oleo.Col = i
    grid_oleo.Text = "Número do Documento"
    grid_oleo.ColWidth(i) = 1000
    grid_oleo.ColAlignment(i) = 0
    i = i + 1
    grid_oleo.Col = i
    grid_oleo.Text = "Produto"
    grid_oleo.ColWidth(i) = 2000
    grid_oleo.ColAlignment(i) = 0
    i = i + 1
    grid_oleo.Col = i
    grid_oleo.Text = "Preço de Custo"
    grid_oleo.ColWidth(i) = 700
    grid_oleo.ColAlignment(i) = 1
    i = i + 1
    grid_oleo.Col = i
    grid_oleo.Text = "Quantidade"
    grid_oleo.ColWidth(i) = 900
    grid_oleo.ColAlignment(i) = 1
    i = i + 1
    grid_oleo.Col = i
    grid_oleo.Text = "Tipo Ent."
    grid_oleo.ColWidth(i) = 400
    grid_oleo.ColAlignment(i) = 2
    i = i + 1
    grid_oleo.Col = i
    grid_oleo.Text = "Data da Digitação"
    grid_oleo.ColWidth(i) = 1000
    grid_oleo.ColAlignment(i) = 2
    i = i + 1
    grid_oleo.Col = i
    grid_oleo.Text = "Observação"
    grid_oleo.ColWidth(i) = 2000
    grid_oleo.ColAlignment(i) = 0
    i = i + 1
    grid_oleo.Col = i
    grid_oleo.Text = "Código do Produto"
    grid_oleo.ColWidth(i) = 700
    grid_oleo.ColAlignment(i) = 1
End Sub
Private Sub txt_numero_documento_GotFocus()
    txt_numero_documento.SelStart = 0
    txt_numero_documento.SelLength = Len(txt_numero_documento)
End Sub
Private Sub txt_numero_documento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_tipo_entrada.SetFocus
    End If
End Sub
Private Sub txt_observacao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub txt_preco_custo_GotFocus()
    txt_preco_custo.SelStart = 0
    txt_preco_custo.SelLength = Len(txt_preco_custo.Text)
End Sub
Private Sub txt_preco_custo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_quantidade.SetFocus
    End If
    Call ValidaValor(KeyAscii)
End Sub
Private Sub txt_preco_custo_LostFocus()
    txt_preco_custo.Text = Format(txt_preco_custo.Text, "###,##0.0000")
End Sub
Private Sub txt_preco_total_GotFocus()
    txt_preco_total.SelStart = 0
    txt_preco_total.SelLength = Len(txt_preco_total.Text)
End Sub
Private Sub txt_preco_total_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
    Call ValidaValor(KeyAscii)
End Sub
Private Sub txt_preco_total_LostFocus()
    txt_preco_total.Text = Format(txt_preco_total.Text, "###,##0.00")
    If fValidaValor(txt_quantidade.Text) > 0 Then
        If fValidaValor(txt_preco_total.Text) > 0 Then
            txt_preco_custo.Text = Format(fValidaValor(txt_preco_total.Text) / fValidaValor(txt_quantidade.Text), "###,##0.0000")
        End If
    End If
End Sub
Private Sub txt_produto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dtcboProduto.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_produto_LostFocus()
    If Val(txt_produto.Text) > 0 And lOpcao > 0 Then
        If Len(txt_produto.Text) > 5 Then
            If Produto.LocalizarCodigoBarra(txt_produto.Text) Then
                txt_produto.Text = Produto.Codigo
            Else
                MsgBox "Codigo de Barra não cadastrado!", vbInformation, "Erro de Leitura!"
                txt_produto.Text = ""
                txt_produto.SetFocus
                Exit Sub
            End If
        End If
        If Produto.LocalizarCodigo(CLng(txt_produto.Text)) Then
            dtcboProduto.BoundText = CLng(txt_produto.Text)
            lbl_unidade.Caption = Produto.Unidade
            txt_preco_custo.Text = Format(Produto.PrecoCusto, "###,##0.0000")
            txt_quantidade.SetFocus
        Else
            MsgBox "Produto não cadastrado.", vbInformation, "Atenção!"
            txt_produto.SetFocus
            Exit Sub
        End If
    End If
End Sub
Function ExisteRegistro() As Boolean
    ExisteRegistro = False
    If EntradaProduto.LocalizarCodigo(g_empresa, CDate(msk_data_entrada.Text), CLng(txt_produto.Text), txt_numero_documento.Text) Then
        MsgBox "Já existe movimento com este produto." & Chr(10) & Chr(10) & "Mude o produto informado.", vbInformation, "Duplicidade de Registro!"
        ExisteRegistro = True
    End If
End Function
Private Sub txt_quantidade_GotFocus()
    txt_quantidade.SelStart = 0
    txt_quantidade.SelLength = Len(txt_quantidade.Text)
End Sub
Private Sub txt_quantidade_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_preco_total.SetFocus
        If cbo_tipo_entrada.ItemData(cbo_tipo_entrada.ListIndex) = 3 Then
            cmd_ok.SetFocus
        End If
    End If
    Call ValidaValorSinal(KeyAscii)
End Sub
Private Sub txt_quantidade_LostFocus()
    txt_quantidade.Text = Format(txt_quantidade.Text, "###,##0.00")
    txt_preco_total.Text = Format(fValidaValor(txt_preco_custo.Text) * fValidaValor(txt_quantidade.Text), "###,##0.00")
    msk_data_digitacao.Text = Format(lDataDigitacao, "dd/mm/yyyy")
End Sub
