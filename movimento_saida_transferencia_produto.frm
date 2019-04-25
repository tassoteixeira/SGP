VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C21-0000-0000-0000-000000000046}#4.0#0"; "TDBG32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form movimento_saida_transferencia_produto 
   Caption         =   "Movimento de Saida de Transferência de Produtos"
   ClientHeight    =   6375
   ClientLeft      =   165
   ClientTop       =   285
   ClientWidth     =   8550
   Icon            =   "movimento_saida_transferencia_produto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "movimento_saida_transferencia_produto.frx":030A
   ScaleHeight     =   6375
   ScaleWidth      =   8550
   Begin VB.CommandButton cmd_sair 
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3720
      Picture         =   "movimento_saida_transferencia_produto.frx":0750
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   5400
      Width           =   795
   End
   Begin VB.CommandButton cmd_pesquisa 
      Caption         =   "&Pesquisa"
      Height          =   855
      Left            =   2820
      Picture         =   "movimento_saida_transferencia_produto.frx":1DE2
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Pesquisa um registro específico."
      Top             =   5400
      Width           =   795
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   1920
      Picture         =   "movimento_saida_transferencia_produto.frx":3254
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Exclui o registro atual."
      Top             =   5400
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   1020
      Picture         =   "movimento_saida_transferencia_produto.frx":48E6
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Altera o registro atual."
      Top             =   5400
      Width           =   795
   End
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   120
      Picture         =   "movimento_saida_transferencia_produto.frx":5DE0
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Cria um novo registro."
      Top             =   5400
      Width           =   795
   End
   Begin VB.Data dta_saida_transferencia_produto 
      Caption         =   "dta_saida_transferencia_produto"
      Connect         =   "Access"
      DatabaseName    =   "Sgp_data.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Saida_Transferencia_Produto"
      Top             =   4920
      Visible         =   0   'False
      Width           =   3795
   End
   Begin VB.Frame frmDados 
      Enabled         =   0   'False
      Height          =   3195
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   8295
      Begin VB.Data dta_empresa 
         Caption         =   "dta_empresa"
         Connect         =   "Access"
         DatabaseName    =   "Sgp_data.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2940
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Empresas"
         Top             =   990
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox txt_observacao 
         Height          =   285
         Left            =   3000
         MaxLength       =   40
         TabIndex        =   21
         Top             =   2760
         Width           =   5175
      End
      Begin VB.TextBox txt_numero_documento 
         Height          =   285
         Left            =   3000
         MaxLength       =   9
         TabIndex        =   4
         Top             =   420
         Width           =   1095
      End
      Begin VB.Data dta_produto 
         Caption         =   "dta_produto"
         Connect         =   "Access"
         DatabaseName    =   "Sgp_data.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2940
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Produto"
         Top             =   1560
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox txt_quantidade 
         Height          =   285
         Left            =   3000
         MaxLength       =   10
         TabIndex        =   15
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox txt_preco_custo 
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   13
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox txt_produto 
         Height          =   285
         Left            =   120
         MaxLength       =   6
         TabIndex        =   8
         Top             =   1560
         Width           =   795
      End
      Begin MSMask.MaskEdBox msk_data_transferencia 
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
      Begin MSDBCtls.DBCombo dbcbo_produto 
         Bindings        =   "movimento_saida_transferencia_produto.frx":7472
         Height          =   315
         Left            =   960
         TabIndex        =   9
         Top             =   1560
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSMask.MaskEdBox msk_data_digitacao 
         Height          =   300
         Left            =   120
         TabIndex        =   19
         Top             =   2760
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
      Begin MSDBCtls.DBCombo dbcbo_empresa 
         Bindings        =   "movimento_saida_transferencia_produto.frx":748C
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   990
         Width           =   5475
         _ExtentX        =   9657
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         Caption         =   "&Empresa"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   5
         Top             =   780
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Observação"
         Height          =   195
         Index           =   9
         Left            =   3000
         TabIndex        =   20
         Top             =   2550
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Data da digitação"
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   18
         Top             =   2550
         Width           =   1815
      End
      Begin VB.Label lbl_total 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   6000
         TabIndex        =   17
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label lbl_unidade 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   6000
         TabIndex        =   11
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Unidade"
         Height          =   195
         Index           =   3
         Left            =   6000
         TabIndex        =   10
         Top             =   1350
         Width           =   675
      End
      Begin VB.Label Label3 
         Caption         =   "Numero do documento"
         Height          =   195
         Index           =   6
         Left            =   3000
         TabIndex        =   3
         Top             =   210
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "&Quantidade"
         Height          =   195
         Index           =   2
         Left            =   3000
         TabIndex        =   14
         Top             =   1950
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "&Preço de Custo"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   1950
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "P&roduto"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   1350
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Total da Entrada"
         Height          =   195
         Index           =   5
         Left            =   6000
         TabIndex        =   16
         Top             =   1950
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Data da transferência"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   1
         Top             =   210
         Width           =   1815
      End
   End
   Begin TrueDBGrid.TDBGrid TDBGrid1 
      Bindings        =   "movimento_saida_transferencia_produto.frx":74A6
      Height          =   1995
      Left            =   120
      OleObjectBlob   =   "movimento_saida_transferencia_produto.frx":74D4
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   3300
      Width           =   8295
   End
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   6240
      TabIndex        =   30
      Top             =   5280
      Width           =   2175
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "movimento_saida_transferencia_produto.frx":A004
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "movimento_saida_transferencia_produto.frx":B586
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "movimento_saida_transferencia_produto.frx":C9F8
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "movimento_saida_transferencia_produto.frx":DEF2
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   6720
      Picture         =   "movimento_saida_transferencia_produto.frx":F3EC
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Confirma o registro atual."
      Top             =   5400
      Width           =   795
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   7620
      Picture         =   "movimento_saida_transferencia_produto.frx":109F6
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Cancela o registro atual."
      Top             =   5400
      Width           =   795
   End
End
Attribute VB_Name = "movimento_saida_transferencia_produto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flag_movimento_saida_transferencia_produto As Integer
Dim lOpcao As String
Dim l_empresa As Integer
Dim l_data As Date
Dim l_codigo_produto As Long
Dim l_numero_documento As Long
Dim l_quantidade As Currency
Dim l_data_digitacao As Date
Dim l_sql As String
Dim l_gravados As Long
Dim l_entrou_empresa As Integer

Private EntradaProduto As New cEntradaProduto
Private Estoque As New cEstoque
Private Produto As New cProduto
Private SaidaTransferenciaProduto As New cSaidaTransferenciaProduto
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
Private Sub AtualizaTDBGrid1()
    Dim i As Integer
    While TDBGrid1.Columns.Count <> 0
        TDBGrid1.Columns.Remove 0
    Wend
    For i = 0 To 8
        TDBGrid1.Columns.Add 0
    Next
    For i = 0 To 7
        TDBGrid1.Columns(i).Visible = True
    Next
    TDBGrid1.Columns(0).DataField = "Data da Transferencia"
    TDBGrid1.Columns(0).NumberFormat = "General Date"
    TDBGrid1.Columns(0).Caption = "Data da Transferência"
    TDBGrid1.Columns(0).Alignment = dbgCenter
    TDBGrid1.Columns(0).HeadAlignment = dbgCenter
    TDBGrid1.Columns(0).Width = 1100
    TDBGrid1.Columns(1).DataField = "Numero do Documento"
    TDBGrid1.Columns(1).Caption = "Número do Documento"
    TDBGrid1.Columns(1).Alignment = dbgRight
    TDBGrid1.Columns(1).HeadAlignment = dbgCenter
    TDBGrid1.Columns(1).Width = 1000
    TDBGrid1.Columns(2).DataField = "Nome"
    TDBGrid1.Columns(2).Caption = "Produto"
    TDBGrid1.Columns(2).Alignment = dbgLeft
    TDBGrid1.Columns(2).HeadAlignment = dbgCenter
    TDBGrid1.Columns(2).Width = 2000
    TDBGrid1.Columns(3).DataField = "Preco de Custo"
    TDBGrid1.Columns(3).NumberFormat = "Currency"
    TDBGrid1.Columns(3).Caption = "Preço de Custo"
    TDBGrid1.Columns(3).Alignment = dbgRight
    TDBGrid1.Columns(3).HeadAlignment = dbgCenter
    TDBGrid1.Columns(3).Width = 700
    TDBGrid1.Columns(4).DataField = "Quantidade"
    TDBGrid1.Columns(4).NumberFormat = "Currency"
    TDBGrid1.Columns(4).Caption = "Quantidade"
    TDBGrid1.Columns(4).Alignment = dbgRight
    TDBGrid1.Columns(4).HeadAlignment = dbgCenter
    TDBGrid1.Columns(4).Width = 900
    TDBGrid1.Columns(5).DataField = "NomeEmpresa"
    TDBGrid1.Columns(5).Caption = "Entrou na Empresa"
    TDBGrid1.Columns(5).Alignment = dbgLeft
    TDBGrid1.Columns(5).HeadAlignment = dbgCenter
    TDBGrid1.Columns(5).Width = 2500
    TDBGrid1.Columns(6).DataField = "Data da Digitacao"
    TDBGrid1.Columns(6).NumberFormat = "General Date"
    TDBGrid1.Columns(6).Caption = "Data da Digitacao"
    TDBGrid1.Columns(6).Alignment = dbgCenter
    TDBGrid1.Columns(6).HeadAlignment = dbgCenter
    TDBGrid1.Columns(6).Width = 1000
    TDBGrid1.Columns(7).DataField = "Observacao"
    TDBGrid1.Columns(7).Caption = "Observação"
    TDBGrid1.Columns(7).Alignment = dbgLeft
    TDBGrid1.Columns(7).HeadAlignment = dbgCenter
    TDBGrid1.Columns(7).Width = 2000
    TDBGrid1.Columns(8).DataField = "Codigo do Produto2"
    TDBGrid1.Columns(8).Caption = "Código do Produto"
    TDBGrid1.Columns(8).Alignment = dbgRight
    TDBGrid1.Columns(8).HeadAlignment = dbgCenter
    TDBGrid1.Columns(8).Width = 700
    If Val(l_numero_documento) = 0 Then
        l_numero_documento = 0
        l_data = g_data_def
    End If
    l_sql = "Select Saida_Transferencia_Produto.[Data da Transferencia], Saida_Transferencia_Produto.[Numero do Documento], Produto.Nome, Saida_Transferencia_Produto.[Preco de Custo], Saida_Transferencia_Produto.Quantidade, Empresas.Nome As NomeEmpresa, Saida_Transferencia_Produto.[Data da Digitacao], Saida_Transferencia_Produto.Observacao, Saida_Transferencia_Produto.[Codigo do Produto2]"
    l_sql = l_sql & " From Saida_Transferencia_Produto, Produto, Empresas"
    l_sql = l_sql & " Where Saida_Transferencia_Produto.Empresa = " & g_empresa
    l_sql = l_sql & " And Saida_Transferencia_Produto.[Data da Transferencia] = #" & CDate(Format(l_data, "mm/dd/yyyy")) & "#"
    l_sql = l_sql & " And Saida_Transferencia_Produto.[Numero do Documento] = " & l_numero_documento
    l_sql = l_sql & " And Produto.Codigo = Saida_Transferencia_Produto.[Codigo do Produto2]"
    l_sql = l_sql & " And Empresas.Codigo = Saida_Transferencia_Produto.[Entrou na Empresa]"
    l_sql = l_sql & " Order by [Data da Transferencia], [Codigo do Produto2], [Numero do Documento]"
    dta_saida_transferencia_produto.RecordSource = l_sql
    dta_saida_transferencia_produto.Refresh
    If dta_saida_transferencia_produto.Recordset.RecordCount > 0 Then
        dta_saida_transferencia_produto.Recordset.MoveLast
    End If
End Sub
Private Sub AtualTabe()
    SaidaTransferenciaProduto.Empresa = g_empresa
    SaidaTransferenciaProduto.DataTransferencia = CDate(msk_data_transferencia.Text)
    SaidaTransferenciaProduto.CodigoProduto2 = CLng(dbcbo_produto.BoundText)
    SaidaTransferenciaProduto.NumeroDocumento = CLng(txt_numero_documento.Text)
    SaidaTransferenciaProduto.EntrounaEmpresa = Val(dbcbo_empresa.BoundText)
    SaidaTransferenciaProduto.PrecoCusto = fValidaValor2(txt_preco_custo.Text)
    SaidaTransferenciaProduto.Quantidade = fValidaValor2(txt_quantidade.Text)
    SaidaTransferenciaProduto.DataDigitacao = CDate(msk_data_digitacao.Text)
    SaidaTransferenciaProduto.Observacao = txt_observacao.Text
    
    EntradaProduto.Empresa = Val(dbcbo_empresa.BoundText)
    EntradaProduto.DataEntrada = CDate(msk_data_transferencia.Text)
    EntradaProduto.CodigoProduto = CLng(dbcbo_produto.BoundText)
    EntradaProduto.NumeroDocumento = txt_numero_documento.Text
    EntradaProduto.TipoEntrada = 4
    EntradaProduto.PrecoCusto = fValidaValor2(txt_preco_custo.Text)
    EntradaProduto.Quantidade = fValidaValor2(txt_quantidade.Text)
    EntradaProduto.TotalCusto = SaidaTransferenciaProduto.PrecoCusto * SaidaTransferenciaProduto.Quantidade
    EntradaProduto.CodigoFornecedor = 0
    EntradaProduto.DataDigitacao = CDate(msk_data_digitacao.Text)
    EntradaProduto.Observacao = txt_observacao.Text
End Sub
Private Sub AtualTela()
    Dim i As Integer
    l_data = SaidaTransferenciaProduto.DataTransferencia
    l_codigo_produto = SaidaTransferenciaProduto.CodigoProduto2
    l_numero_documento = SaidaTransferenciaProduto.NumeroDocumento
    l_quantidade = SaidaTransferenciaProduto.Quantidade
    l_entrou_empresa = SaidaTransferenciaProduto.EntrounaEmpresa
    msk_data_transferencia.Text = Format(SaidaTransferenciaProduto.DataTransferencia, "dd/mm/yyyy")
    txt_numero_documento.Text = SaidaTransferenciaProduto.NumeroDocumento
    dbcbo_empresa.BoundText = ""
    dbcbo_empresa.BoundText = SaidaTransferenciaProduto.EntrounaEmpresa
    dbcbo_produto.BoundText = ""
    lbl_unidade.Caption = ""
    If Produto.LocalizarCodigo(SaidaTransferenciaProduto.CodigoProduto2) Then
        txt_produto.Text = SaidaTransferenciaProduto.CodigoProduto2
        dbcbo_produto.BoundText = SaidaTransferenciaProduto.CodigoProduto2
        lbl_unidade.Caption = Produto.Unidade
    End If
    txt_preco_custo.Text = Format(SaidaTransferenciaProduto.PrecoCusto, "###,##0.00")
    txt_quantidade.Text = Format(SaidaTransferenciaProduto.Quantidade, "###,##0.00")
    lbl_total.Caption = Format(SaidaTransferenciaProduto.PrecoCusto * SaidaTransferenciaProduto.Quantidade, "###,##0.00")
    msk_data_digitacao.Text = Format(SaidaTransferenciaProduto.DataDigitacao, "dd/mm/yyyy")
    txt_observacao.Text = SaidaTransferenciaProduto.Observacao
    If Not EntradaProduto.LocalizarCodigo(l_entrou_empresa, l_data, l_codigo_produto, CStr(l_numero_documento)) Then
        MsgBox "Não foi possível localizar registro de entrada.", vbInformation, "Erro de Integridade!"
    End If
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
    Set EntradaProduto = Nothing
    Set Estoque = Nothing
    Set Produto = Nothing
    Set SaidaTransferenciaProduto = Nothing
End Sub
Private Sub cmd_alterar_Click()
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
    If SaidaTransferenciaProduto.LocalizarAnterior Then
        AtualTela
    Else
        MsgBox "Início de Arquivo.", vbExclamation, "Atenção!"
        cmd_proximo.SetFocus
    End If
End Sub
Private Sub cmd_cancelar_Click()
    LimpaTela
    If SaidaTransferenciaProduto.LocalizarCodigo(g_empresa, l_data, l_codigo_produto, l_numero_documento) Then
        AtualTela
        AtivaBotoes
        AtualizaTDBGrid1
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
Private Sub LimpaTela()
    If l_gravados = 0 Then
        msk_data_transferencia.Text = "__/__/____"
        txt_numero_documento.Text = ""
        dbcbo_empresa.BoundText = ""
    End If
    txt_produto.Text = ""
    dbcbo_produto.BoundText = ""
    lbl_unidade.Caption = ""
    txt_preco_custo.Text = ""
    txt_quantidade.Text = ""
    lbl_total.Caption = ""
    msk_data_digitacao.Text = "__/__/____"
    txt_observacao.Text = "Transferência de Produto"
End Sub
Private Sub cmd_excluir_Click()
    If Val(txt_produto.Text) > 0 Then
        If (MsgBox("Deseja excluir este registro?", 4 + 32 + 256, "Exclusão de Registro!")) = 6 Then
            lOpcao = 3
            Call ExcluiSaidaTransferenciaProduto(l_codigo_produto, l_quantidade)
            Call ExcluiEntradaProduto(l_entrou_empresa, l_codigo_produto, l_quantidade)
            If SaidaTransferenciaProduto.Excluir(g_empresa, CDate(msk_data_transferencia.Text), CLng(txt_produto.Text), CLng(txt_numero_documento.Text)) Then
                If Not EntradaProduto.Excluir(l_entrou_empresa, CDate(msk_data_transferencia.Text), CLng(txt_produto.Text), txt_numero_documento.Text) Then
                    MsgBox "Não foi possível excluir registro de entrada.", vbInformation, "Erro de Integridade!"
                End If
            Else
                MsgBox "Não foi possível excluir registro de transferência.", vbInformation, "Erro de Integridade!"
            End If
            If SaidaTransferenciaProduto.LocalizarUltimo(g_empresa) Then
                AtualTela
                DesativaBotoes
                cmd_novo.Enabled = True
                cmd_sair.Enabled = True
                cmd_novo.SetFocus
            End If
            lOpcao = 0
            AtualizaTDBGrid1
        End If
    End If
End Sub
Private Sub cmd_novo_Click()
    LimpaTela
    Inclui
    frmDados.Enabled = True
    If l_gravados = 0 Then
        msk_data_transferencia.SetFocus
    Else
        txt_produto.SetFocus
    End If
End Sub
Private Sub cmd_ok_Click()
    On Error GoTo FileError
    If ValidaCampos Then
        AtivaBotoes
        If lOpcao = 1 Then
            AtualTabe
            If SaidaTransferenciaProduto.Incluir Then
                l_data = CDate(msk_data_transferencia.Text)
                l_codigo_produto = CLng(dbcbo_produto.BoundText)
                l_numero_documento = CLng(txt_numero_documento.Text)
                l_quantidade = fValidaValor2(txt_quantidade.Text)
                l_entrou_empresa = Val(dbcbo_empresa.BoundText)
                If Not EntradaProduto.Incluir Then
                    MsgBox "Não foi possível incluir registro de entrada.", vbInformation, "Erro de Integridade!"
                End If
                Call IncluiSaidaTransferenciaProduto(l_codigo_produto, l_quantidade)
                Call IncluiEntradaProduto(l_entrou_empresa, l_codigo_produto, l_quantidade)
                l_gravados = l_gravados + 1
            Else
                MsgBox "Não foi possível incluir registro de transferência.", vbInformation, "Erro de Integridade!"
            End If
        ElseIf lOpcao = 2 Then
            Call ExcluiSaidaTransferenciaProduto(l_codigo_produto, l_quantidade)
            Call ExcluiEntradaProduto(l_entrou_empresa, l_codigo_produto, l_quantidade)
            AtualTabe
            If SaidaTransferenciaProduto.Alterar(g_empresa, l_data, l_codigo_produto, l_numero_documento) Then
                If Not EntradaProduto.Alterar(l_entrou_empresa, l_data, l_codigo_produto, CStr(l_numero_documento)) Then
                    MsgBox "Não foi possível alterar registro de entrada.", vbInformation, "Erro de Integridade!"
                End If
                l_data = CDate(msk_data_transferencia.Text)
                l_codigo_produto = CLng(dbcbo_produto.BoundText)
                l_numero_documento = CLng(txt_numero_documento.Text)
                l_quantidade = fValidaValor2(txt_quantidade.Text)
                l_entrou_empresa = Val(dbcbo_empresa.BoundText)
                Call IncluiSaidaTransferenciaProduto(l_codigo_produto, l_quantidade)
                Call IncluiEntradaProduto(l_entrou_empresa, l_codigo_produto, l_quantidade)
            Else
                MsgBox "Não foi possível alterar registro de transferência.", vbInformation, "Erro de Integridade!"
            End If
        End If
        AtualizaTDBGrid1
        If SaidaTransferenciaProduto.LocalizarCodigo(g_empresa, l_data, l_codigo_produto, l_numero_documento) Then
            AtualTela
        Else
            MsgBox "Não foi possível localizar registro de transferencia.", vbInformation, "Erro de Integridade!"
        End If
        cmd_novo.SetFocus
        If lOpcao = 1 Then
            lOpcao = 0
            cmd_novo_Click
        Else
            lOpcao = 0
        End If
    End If
    Exit Sub
FileError:
    MsgBox "Erro na atualização de dados", vbInformation, "Erro Desconhecido"
    Exit Sub
End Sub
Function ValidaCampos() As Boolean
    ValidaCampos = False
    If Not IsDate(msk_data_transferencia) Then
        MsgBox "Informe a data da transferencia.", vbInformation, "Atenção!"
        msk_data_transferencia.SetFocus
    ElseIf Not CLng(txt_numero_documento) > 0 Then
        MsgBox "Informe o número do documento.", vbInformation, "Atenção!"
        txt_numero_documento.SetFocus
    ElseIf dbcbo_empresa.BoundText = "" Then
        MsgBox "Escolha a empresa.", vbInformation, "Atenção!"
        dbcbo_empresa.SetFocus
    ElseIf Val(dbcbo_empresa.BoundText) = g_empresa Then
        MsgBox "A empresa não pode ser a mesma!" & Chr(10) & "Escolha outra empresa.", vbInformation, "Atenção!"
        dbcbo_empresa.SetFocus
    ElseIf dbcbo_produto.BoundText = "" Then
        MsgBox "Escolha o produto.", vbInformation, "Atenção!"
        dbcbo_produto.SetFocus
    ElseIf Not fValidaValor2(txt_preco_custo) > 0 Then
        MsgBox "Informe o preço de custo.", vbInformation, "Atenção!"
        txt_preco_custo.SetFocus
    ElseIf Not fValidaValor2(txt_quantidade) > 0 Then
        MsgBox "Informe a quantidade.", vbInformation, "Atenção!"
        txt_quantidade.SetFocus
    ElseIf Not IsDate(msk_data_digitacao) Then
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
        If SaidaTransferenciaProduto.Empresa < g_cfg_empresa_i Or SaidaTransferenciaProduto.Empresa > g_cfg_empresa_f Then
            x_flag = False
        ElseIf SaidaTransferenciaProduto.DataTransferencia < g_cfg_data_i Or SaidaTransferenciaProduto.DataTransferencia > g_cfg_data_f Then
            x_flag = False
        End If
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
    If msk_data_transferencia < g_cfg_data_i Or msk_data_transferencia > g_cfg_data_f Then
        MsgBox "A data da entrada deve estar entre " & Format(g_cfg_data_i, "dd/mm/yyyy") & " a " & Format(g_cfg_data_f, "dd/mm/yyyy") & ".", vbInformation, "Digitação Não Autorizada!"
        msk_data_transferencia.SetFocus
    Else
        VerificaLiberacaoDigitacao2 = True
    End If
End Function
Private Sub cmd_pesquisa_Click()
    consulta_movimento_saida_transferencia_produto.Show 1
    If Len(g_string) > 0 Then
        l_data = RetiraGString(1)
        l_codigo_produto = RetiraGString(2)
        l_numero_documento = RetiraGString(3)
        If SaidaTransferenciaProduto.LocalizarCodigo(g_empresa, l_data, l_codigo_produto, l_numero_documento) Then
            AtualTela
            AtualizaTDBGrid1
        Else
            MsgBox "Não foi possível localizar registro de transferencia.", vbInformation, "Erro de Integridade!"
        End If
    End If
End Sub
Private Sub cmd_primeiro_Click()
    If SaidaTransferenciaProduto.LocalizarPrimeiro Then
        AtualTela
        cmd_proximo.SetFocus
    Else
        MsgBox "Não há registro nesta empresa.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    If SaidaTransferenciaProduto.LocalizarProximo Then
        AtualTela
    Else
        MsgBox "Fim de Arquivo.", vbExclamation, "Atenção!"
        cmd_anterior.SetFocus
    End If
End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub cmd_ultimo_Click()
    If SaidaTransferenciaProduto.LocalizarUltimo(g_empresa) Then
        AtualTela
        cmd_anterior.SetFocus
    Else
        MsgBox "Não há registro nesta empresa.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub dbcbo_empresa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        dbcbo_produto.SetFocus
    End If
End Sub
Private Sub dbcbo_produto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        txt_quantidade.SetFocus
    End If
End Sub
Private Sub dbcbo_produto_LostFocus()
    If dbcbo_produto.BoundText <> "" And lOpcao > 0 Then
        txt_produto = dbcbo_produto.BoundText
        If lOpcao = 1 Then
            If ExisteRegistro Or ExisteRegistroEntrada Then
                txt_produto = ""
                dbcbo_produto.BoundText = ""
                txt_produto.SetFocus
                Exit Sub
            End If
        End If
        txt_produto_LostFocus
        txt_quantidade.SetFocus
    End If
End Sub
Private Sub Form_Activate()
    If g_empresa <> l_empresa Then
        flag_movimento_saida_transferencia_produto = 0
    End If
    If flag_movimento_saida_transferencia_produto = 0 Then
        l_gravados = 0
        lOpcao = 0
        l_empresa = g_empresa
        DesativaBotoes
        If SaidaTransferenciaProduto.LocalizarUltimo(g_empresa) Then
            AtualTela
            AtivaBotoes
            AtualizaTDBGrid1
        Else
            cmd_novo.Enabled = True
            cmd_sair.Enabled = True
            AtualizaTDBGrid1
        End If
        cmd_novo.SetFocus
    Else
        flag_movimento_saida_transferencia_produto = 0
    End If
    Screen.MousePointer = 1
End Sub
Private Sub Form_Deactivate()
    flag_movimento_saida_transferencia_produto = 1
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
    
    dta_empresa.RecordSource = "Select * From Empresas Where Codigo <> " & g_empresa & " Order By Codigo"
    dta_empresa.Refresh
    dta_produto.RecordSource = "Select * From Produto Order By Nome"
    dta_produto.Refresh
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_digitacao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_observacao.SetFocus
    End If
End Sub
Private Sub msk_data_digitacao_LostFocus()
    If IsDate(msk_data_digitacao) Then
        l_data_digitacao = msk_data_digitacao
    End If
End Sub
Private Sub msk_data_transferencia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_numero_documento.SetFocus
    End If
End Sub
Private Sub msk_data_transferencia_GotFocus()
    If Not IsDate(msk_data_transferencia) Then
        msk_data_transferencia = Format(g_data_def, "dd/mm/yyyy")
        l_data_digitacao = g_data_def
    End If
End Sub
Private Sub TDBGrid1_DblClick()
    MarcaCelulas
End Sub
Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        MarcaCelulas
    ElseIf KeyCode = 98 Then
        KeyCode = 40
    ElseIf KeyCode = 99 Then
        KeyCode = 34
    ElseIf KeyCode = 104 Then
        KeyCode = 38
    ElseIf KeyCode = 105 Then
        KeyCode = 33
    End If
End Sub
Private Sub MarcaCelulas()
    l_data = TDBGrid1.Columns(0).Text
    l_codigo_produto = TDBGrid1.Columns(8).Text
    l_numero_documento = TDBGrid1.Columns(1).Text
    If SaidaTransferenciaProduto.LocalizarCodigo(g_empresa, l_data, l_codigo_produto, l_numero_documento) Then
        AtualTela
    Else
        MsgBox "Não foi possível localizar registro de transferencia.", vbInformation, "Erro de Integridade!"
    End If
    cmd_alterar.SetFocus
End Sub
Private Sub TDBGrid1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
        KeyAscii = 0
        MarcaCelulas
    End If
End Sub
Private Sub txt_numero_documento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dbcbo_empresa.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_observacao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub txt_preco_custo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_quantidade.SetFocus
    End If
End Sub
Private Sub txt_preco_custo_LostFocus()
    txt_preco_custo = Format(txt_preco_custo, "###,##0.00")
End Sub
Private Sub txt_produto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dbcbo_produto.SetFocus
    End If
End Sub
Private Sub txt_produto_LostFocus()
    If Val(txt_produto.Text) > 0 And lOpcao > 0 Then
        If Produto.LocalizarCodigo(CLng(txt_produto.Text)) Then
            dbcbo_produto.BoundText = CLng(txt_produto.Text)
            lbl_unidade.Caption = Produto.Unidade
            txt_preco_custo.Text = Format(Produto.PrecoCusto, "###,##0.00")
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
    If SaidaTransferenciaProduto.LocalizarCodigo(g_empresa, CDate(msk_data_transferencia.Text), CLng(txt_produto.Text), Val(txt_numero_documento.Text)) Then
        MsgBox "Já existe movimento de transferência com este produto." & Chr(10) & Chr(10) & "Mude o produto informado.", vbInformation, "Duplicidade de Registro!"
        ExisteRegistro = True
    End If
End Function
Function ExisteRegistroEntrada() As Boolean
    ExisteRegistroEntrada = False
    If EntradaProduto.LocalizarCodigo(Val(dbcbo_empresa.BoundText), CDate(msk_data_transferencia.Text), CLng(txt_produto.Text), txt_numero_documento.Text) Then
        MsgBox "Já existe movimento de entrada com este produto." & Chr(10) & Chr(10) & "Mude o produto informado.", vbInformation, "Duplicidade de Registro!"
        ExisteRegistroEntrada = True
    End If
End Function
Private Sub txt_quantidade_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub txt_quantidade_LostFocus()
    txt_quantidade = Format(txt_quantidade, "###,##0.00")
    lbl_total = Format(fValidaValor2(txt_preco_custo) * fValidaValor2(txt_quantidade), "###,##0.00")
    msk_data_digitacao = Format(l_data_digitacao, "dd/mm/yyyy")
End Sub
Private Sub IncluiEntradaProduto(x_empresa As Integer, x_codigo2 As Long, x_quantidade As Currency)
    If Estoque.LocalizarCodigo(x_empresa, x_codigo2) Then
        Estoque.Quantidade = Estoque.Quantidade + x_quantidade
        If Not Estoque.Alterar(x_empresa, x_codigo2) Then
            MsgBox "Não foi possível alterar registro de estoque.", vbInformation, "Erro de Integridade!"
        End If
    End If
End Sub
Private Sub IncluiSaidaTransferenciaProduto(x_codigo2 As Long, x_quantidade As Currency)
    If Estoque.LocalizarCodigo(g_empresa, x_codigo2) Then
        Estoque.Quantidade = Estoque.Quantidade - x_quantidade
        If Not Estoque.Alterar(g_empresa, x_codigo2) Then
            MsgBox "Não foi possível alterar registro de estoque.", vbInformation, "Erro de Integridade!"
        End If
    End If
End Sub
Private Sub ExcluiEntradaProduto(x_empresa As Integer, x_codigo2 As Long, x_quantidade As Currency)
    If Estoque.LocalizarCodigo(x_empresa, x_codigo2) Then
        Estoque.Quantidade = Estoque.Quantidade - x_quantidade
        If Not Estoque.Alterar(x_empresa, x_codigo2) Then
            MsgBox "Não foi possível alterar registro de estoque.", vbInformation, "Erro de Integridade!"
        End If
    End If
End Sub
Private Sub ExcluiSaidaTransferenciaProduto(x_codigo2 As Long, x_quantidade As Currency)
    If Estoque.LocalizarCodigo(g_empresa, x_codigo2) Then
        Estoque.Quantidade = Estoque.Quantidade + x_quantidade
        If Not Estoque.Alterar(g_empresa, x_codigo2) Then
            MsgBox "Não foi possível alterar registro de estoque.", vbInformation, "Erro de Integridade!"
        End If
    End If
End Sub
