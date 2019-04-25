VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form baixa_nota_abastecimento 
   Caption         =   "Baixa de Notas de Abastecimento"
   ClientHeight    =   7215
   ClientLeft      =   3615
   ClientTop       =   585
   ClientWidth     =   7890
   Icon            =   "baixa_nota_abastecimento.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   Picture         =   "baixa_nota_abastecimento.frx":030A
   ScaleHeight     =   7215
   ScaleWidth      =   7890
   Begin VB.Frame frmDados 
      Height          =   2655
      Left            =   120
      TabIndex        =   4
      Top             =   3480
      Width           =   7635
      Begin VB.TextBox txt_valor_baixado 
         Height          =   285
         Left            =   6420
         MaxLength       =   10
         TabIndex        =   25
         Top             =   2220
         Width           =   1095
      End
      Begin MSMask.MaskEdBox msk_data_baixa 
         Height          =   300
         Left            =   2100
         TabIndex        =   23
         Top             =   2220
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
      Begin VB.Label lblOrdem 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   7200
         TabIndex        =   37
         Top             =   960
         Width           =   315
      End
      Begin VB.Label Label3 
         Caption         =   "Data do abastecimento"
         Height          =   315
         Index           =   4
         Left            =   180
         TabIndex        =   7
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Preço total"
         Height          =   315
         Index           =   5
         Left            =   5160
         TabIndex        =   20
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Período"
         Height          =   315
         Index           =   6
         Left            =   5040
         TabIndex        =   9
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo do movimento"
         Height          =   315
         Index           =   7
         Left            =   180
         TabIndex        =   11
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Número da nota"
         Height          =   315
         Index           =   10
         Left            =   4680
         TabIndex        =   13
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Produto"
         Height          =   315
         Index           =   0
         Left            =   180
         TabIndex        =   15
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Quantidade"
         Height          =   315
         Index           =   2
         Left            =   180
         TabIndex        =   18
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "&Data da baixa"
         Height          =   315
         Index           =   1
         Left            =   180
         TabIndex        =   22
         Top             =   2220
         Width           =   1815
      End
      Begin VB.Line Line1 
         X1              =   7620
         X2              =   0
         Y1              =   2100
         Y2              =   2100
      End
      Begin VB.Label Label3 
         Caption         =   "&Valor Baixado"
         Height          =   315
         Index           =   3
         Left            =   4500
         TabIndex        =   24
         Top             =   2220
         Width           =   1815
      End
      Begin VB.Label lbl_data_abastecimento 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2100
         TabIndex        =   8
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lbl_periodo 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   6420
         TabIndex        =   10
         Top             =   600
         Width           =   315
      End
      Begin VB.Label lbl_tipo_movimento 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2100
         TabIndex        =   12
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label lbl_nome_produto 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   3000
         TabIndex        =   17
         Top             =   1320
         Width           =   4515
      End
      Begin VB.Label lbl_codigo_produto 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2100
         TabIndex        =   16
         Top             =   1320
         Width           =   795
      End
      Begin VB.Label lbl_numero_nota 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   6060
         TabIndex        =   14
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lbl_total 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   6420
         TabIndex        =   21
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label lbl_quantidade 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2100
         TabIndex        =   19
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label lbl_cliente 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2100
         TabIndex        =   6
         Top             =   240
         Width           =   5415
      End
      Begin VB.Label Label3 
         Caption         =   "Cliente"
         Height          =   315
         Index           =   9
         Left            =   180
         TabIndex        =   5
         Top             =   240
         Width           =   1815
      End
   End
   Begin MSAdodcLib.Adodc adodcCliente 
      Height          =   330
      Left            =   4320
      Top             =   120
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
      Caption         =   "adodcCliente"
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
   Begin VB.CommandButton cmd_extornar 
      Caption         =   "&Extornar"
      Height          =   855
      Left            =   1020
      Picture         =   "baixa_nota_abastecimento.frx":0750
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Extorna o registro atual."
      Top             =   6240
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   2820
      Picture         =   "baixa_nota_abastecimento.frx":1A2A
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   6240
      Width           =   795
   End
   Begin VB.CommandButton cmd_pesquisa 
      Caption         =   "&Pesquisa"
      Height          =   855
      Left            =   1920
      Picture         =   "baixa_nota_abastecimento.frx":30BC
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Pesquisa um registro específico."
      Top             =   6240
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   120
      Picture         =   "baixa_nota_abastecimento.frx":452E
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Altera o registro atual."
      Top             =   6240
      Width           =   795
   End
   Begin VB.TextBox txt_cliente 
      Height          =   285
      Left            =   2040
      MaxLength       =   6
      TabIndex        =   1
      Top             =   120
      Width           =   795
   End
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   5580
      TabIndex        =   32
      Top             =   6120
      Width           =   2175
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "baixa_nota_abastecimento.frx":5A28
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "baixa_nota_abastecimento.frx":6FAA
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "baixa_nota_abastecimento.frx":841C
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "baixa_nota_abastecimento.frx":9916
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   6960
      Picture         =   "baixa_nota_abastecimento.frx":AE10
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Cancela o registro atual."
      Top             =   6240
      Width           =   795
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   6060
      Picture         =   "baixa_nota_abastecimento.frx":C30A
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Confirma o registro atual."
      Top             =   6240
      Width           =   795
   End
   Begin MSDataListLib.DataCombo dtcboCliente 
      Bindings        =   "baixa_nota_abastecimento.frx":D914
      Height          =   315
      Left            =   2940
      TabIndex        =   2
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "Razao Social"
      BoundColumn     =   "Codigo"
      Text            =   "dtcboCliente"
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid 
      Height          =   2955
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   5212
      _Version        =   393216
      Cols            =   9
      FixedCols       =   0
   End
   Begin VB.Label Label3 
      Caption         =   "C&liente"
      Height          =   315
      Index           =   8
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "baixa_nota_abastecimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flag_baixa_nota_abastecimento As Integer
Dim lOpcao As String
Dim lEmpresa As Integer
Dim lEmpresa_baixa As Integer
Dim lDataAbastecimento As Date
Dim lDataPagamento As Date
Dim lPeriodo As String
Dim lTipoMovimento As String
Dim lNumeroNota As Long
Dim lOrdem As Integer
Dim lCodigoCliente As Long
Dim lCodigoProduto As Long
Dim lSQL As String
Dim lNumeroMovimentoCaixa As Long

Private BaixaNotaAbastecimento As New cBaixaNotaAbastecimento
Private Cliente As New cCliente
Private ConfiguracaoDiversa As New cConfiguracaoDiversa
Private IntegracaoCaixa As New cIntegracaoCaixa
Private MovCaixa As New cMovimentoCaixa
Private MovNotaAbastecimento As New cMovimentoNotaAbastecimento
Private Produto As New cProduto
Private Sub AtivaBotoes()
    cmd_alterar.Enabled = True
    cmd_extornar.Enabled = True
    cmd_pesquisa.Enabled = True
    cmd_sair.Enabled = True
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
    frm_move.Visible = True
End Sub
Private Sub Inclui()
    DesativaBotoes
    frmDados.Enabled = True
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
End Sub
Function IncluiMovimentoCaixa() As Boolean
    Dim xComplemento As String
    IncluiMovimentoCaixa = False
    lNumeroMovimentoCaixa = 0
    
    If IntegracaoCaixa.LocalizarNome(g_empresa, "DUPLICATAS A RECEBER") Then
        xComplemento = "TM:" & MovNotaAbastecimento.TipoMovimento & " P:" & MovNotaAbastecimento.Periodo & " " & Cliente.RazaoSocial
        MovCaixa.Empresa = g_empresa
        MovCaixa.Data = Format(msk_data_baixa.Text, "dd/mm/yyyy")
        MovCaixa.NumeroMovimento = 1
        MovCaixa.Valor = MovNotaAbastecimento.ValorTotal
        MovCaixa.NumeroDocumento = MovNotaAbastecimento.NumeroNota
        MovCaixa.CodigoHistorico = IntegracaoCaixa.HistoricoPadrao
        MovCaixa.Complemento = xComplemento
        MovCaixa.NumeroContaDebito = IntegracaoCaixa.ContaDebito
        MovCaixa.NumeroContaCredito = IntegracaoCaixa.ContaCredito
        MovCaixa.TipoMovimento = 2
        MovCaixa.FluxoCaixa = True
        MovCaixa.CodigoUsuario = g_usuario
        If MovCaixa.Incluir > 0 Then
            IncluiMovimentoCaixa = True
            lNumeroMovimentoCaixa = MovCaixa.NumeroMovimento
        Else
            MsgBox "Não foi integrado no caixa o valor=" & MovNotaAbastecimento.ValorTotal, vbInformation, "Erro de Integridade"
        End If
    Else
        MsgBox "Não existe a integração=" & "DUPLICATAS A RECEBER" & ".", vbInformation, "Registro Inexistente"
    End If
End Function
Private Sub AtualizaMSFlexGrid()
    Dim i As Integer
    Dim rsTabela As adodb.Recordset
    
    LimpaMSFlexGrid
    If Val(lTipoMovimento) = 0 Then
        lPeriodo = "1"
        lTipoMovimento = "1"
    End If
    lSQL = "SELECT Movimento_Nota_Abastecimento.[Data do Abastecimento], Movimento_Nota_Abastecimento.Periodo, Movimento_Nota_Abastecimento.[Tipo do Movimento], Movimento_Nota_Abastecimento.[Numero da Nota], Movimento_Nota_Abastecimento.Ordem, Produto.Nome, Movimento_Nota_Abastecimento.[Valor Total], Movimento_Nota_Abastecimento.[Codigo do Produto2], Movimento_Nota_Abastecimento.Empresa"
    lSQL = lSQL & " FROM Movimento_Nota_Abastecimento, Cliente, Produto"
    lSQL = lSQL & " WHERE Cliente.Codigo = Movimento_Nota_Abastecimento.[Codigo do Cliente]"
    lSQL = lSQL & " AND Produto.Codigo = Movimento_Nota_Abastecimento.[Codigo do Produto2]"
    lSQL = lSQL & " AND Movimento_Nota_Abastecimento.[Codigo do Cliente] = " & Val(txt_cliente)
    lSQL = lSQL & " ORDER BY Movimento_Nota_Abastecimento.[Data do Abastecimento], Movimento_Nota_Abastecimento.[Numero da Nota], Movimento_Nota_Abastecimento.Ordem, Movimento_Nota_Abastecimento.[Codigo do Produto2]"
    'Abre RecordSet
    Set rsTabela = New adodb.Recordset
    Set rsTabela = Conectar.RsConexao(lSQL)
    'Verifica movimento
    i = 0
    If rsTabela.RecordCount > 0 Then
        Do Until rsTabela.EOF
            MSFlexGrid.Rows = MSFlexGrid.Rows + 1
            i = i + 1
            MSFlexGrid.Row = i
            MSFlexGrid.Col = 0
            MSFlexGrid.Text = rsTabela![Data do Abastecimento]
            MSFlexGrid.Col = 1
            MSFlexGrid.Text = rsTabela!Periodo
            MSFlexGrid.Col = 2
            MSFlexGrid.Text = rsTabela![Tipo do Movimento]
            MSFlexGrid.Col = 3
            MSFlexGrid.Text = rsTabela![Numero da Nota]
            MSFlexGrid.Col = 4
            MSFlexGrid.Text = rsTabela!Ordem
            MSFlexGrid.Col = 5
            MSFlexGrid.Text = rsTabela!Nome
            MSFlexGrid.Col = 6
            MSFlexGrid.Text = rsTabela![Valor Total]
            MSFlexGrid.Col = 7
            MSFlexGrid.Text = rsTabela![Codigo do Produto2]
            MSFlexGrid.Col = 8
            MSFlexGrid.Text = rsTabela!Empresa
            rsTabela.MoveNext
        Loop
    End If
    If rsTabela.State = 1 Then
        rsTabela.Close
    End If
End Sub
Private Sub AtualTabe()
    BaixaNotaAbastecimento.CodigoCliente = txt_cliente.Text
    BaixaNotaAbastecimento.DataPagamento = Format(msk_data_baixa.Text, "dd/mm/yyyy")
    BaixaNotaAbastecimento.CodigoProduto2 = Val(lbl_codigo_produto.Caption)
    BaixaNotaAbastecimento.NumeroNota = Val(lbl_numero_nota.Caption)
    BaixaNotaAbastecimento.Ordem = Val(lblOrdem.Caption)
    BaixaNotaAbastecimento.Empresa = MovNotaAbastecimento.Empresa
    BaixaNotaAbastecimento.Periodo = lbl_periodo.Caption
    BaixaNotaAbastecimento.DataAbastecimento = Format(lbl_data_abastecimento.Caption, "dd/mm/yyyy")
    BaixaNotaAbastecimento.Quantidade = fValidaValor2(lbl_quantidade.Caption)
    BaixaNotaAbastecimento.ValorUnitario = MovNotaAbastecimento.ValorUnitario
    BaixaNotaAbastecimento.ValorTotal = fValidaValor2(lbl_total.Caption)
    BaixaNotaAbastecimento.CodigoConveniado = MovNotaAbastecimento.CodigoConveniado
    BaixaNotaAbastecimento.TipoMovimento = Mid(lbl_tipo_movimento.Caption, 1, 1)
    BaixaNotaAbastecimento.PlacaLetra = MovNotaAbastecimento.PlacaLetra
    BaixaNotaAbastecimento.PlacaNumero = MovNotaAbastecimento.PlacaNumero
    BaixaNotaAbastecimento.Historico = MovNotaAbastecimento.Historico
    BaixaNotaAbastecimento.ValorPago = fValidaValor2(txt_valor_baixado.Text)
    BaixaNotaAbastecimento.ValorDescontoUnitario = MovNotaAbastecimento.ValorDescontoUnitario
    BaixaNotaAbastecimento.NumeroMovimentoCaixa = MovNotaAbastecimento.NumeroMovimentoCaixa
    BaixaNotaAbastecimento.NumeroMovimentoCaixaBaixa = lNumeroMovimentoCaixa
    BaixaNotaAbastecimento.BaixadoPelaDuplicata = MovNotaAbastecimento.BaixadoPelaDuplicata
    BaixaNotaAbastecimento.Origem = MovNotaAbastecimento.Origem
    BaixaNotaAbastecimento.DataBaixa = CDate(msk_data_baixa.Text)
    BaixaNotaAbastecimento.NumeroCupom = MovNotaAbastecimento.NumeroCupom
    BaixaNotaAbastecimento.DataConferencia = MovNotaAbastecimento.DataConferencia
    BaixaNotaAbastecimento.NumeroDuplicata = 0
    BaixaNotaAbastecimento.KM = MovNotaAbastecimento.KM
    lEmpresa_baixa = BaixaNotaAbastecimento.Empresa
    lDataAbastecimento = BaixaNotaAbastecimento.DataAbastecimento
    lDataPagamento = BaixaNotaAbastecimento.DataPagamento
    lPeriodo = BaixaNotaAbastecimento.Periodo
    lNumeroNota = BaixaNotaAbastecimento.NumeroNota
    lOrdem = BaixaNotaAbastecimento.Ordem
    lCodigoCliente = BaixaNotaAbastecimento.CodigoCliente
    lCodigoProduto = BaixaNotaAbastecimento.CodigoProduto2
    lTipoMovimento = BaixaNotaAbastecimento.TipoMovimento
End Sub
Private Sub AtualizaTabelaNota()
    MovNotaAbastecimento.Empresa = BaixaNotaAbastecimento.Empresa
    MovNotaAbastecimento.CodigoCliente = BaixaNotaAbastecimento.CodigoCliente
    MovNotaAbastecimento.DataAbastecimento = BaixaNotaAbastecimento.DataAbastecimento
    MovNotaAbastecimento.NumeroNota = BaixaNotaAbastecimento.NumeroNota
    MovNotaAbastecimento.Ordem = BaixaNotaAbastecimento.Ordem
    MovNotaAbastecimento.CodigoProduto2 = BaixaNotaAbastecimento.CodigoProduto2
    MovNotaAbastecimento.Periodo = BaixaNotaAbastecimento.Periodo
    MovNotaAbastecimento.Quantidade = BaixaNotaAbastecimento.Quantidade
    MovNotaAbastecimento.ValorUnitario = BaixaNotaAbastecimento.ValorUnitario
    MovNotaAbastecimento.ValorTotal = BaixaNotaAbastecimento.ValorTotal
    MovNotaAbastecimento.CodigoConveniado = BaixaNotaAbastecimento.CodigoConveniado
    MovNotaAbastecimento.TipoMovimento = BaixaNotaAbastecimento.TipoMovimento
    MovNotaAbastecimento.PlacaLetra = BaixaNotaAbastecimento.PlacaLetra
    MovNotaAbastecimento.PlacaNumero = BaixaNotaAbastecimento.PlacaNumero
    MovNotaAbastecimento.Historico = BaixaNotaAbastecimento.Historico
    MovNotaAbastecimento.ValorDescontoUnitario = BaixaNotaAbastecimento.ValorDescontoUnitario
    MovNotaAbastecimento.NumeroMovimentoCaixa = BaixaNotaAbastecimento.NumeroMovimentoCaixa
    MovNotaAbastecimento.BaixadoPelaDuplicata = BaixaNotaAbastecimento.BaixadoPelaDuplicata
    MovNotaAbastecimento.NumeroIlha = BaixaNotaAbastecimento.NumeroIlha
    MovNotaAbastecimento.Origem = BaixaNotaAbastecimento.Origem
    MovNotaAbastecimento.NumeroCupom = BaixaNotaAbastecimento.NumeroCupom
    MovNotaAbastecimento.DataConferencia = BaixaNotaAbastecimento.DataConferencia
    MovNotaAbastecimento.KM = BaixaNotaAbastecimento.KM
End Sub
Private Sub Atualtabe2()
    BaixaNotaAbastecimento.DataPagamento = Format(msk_data_baixa.Text, "dd/mm/yyyy")
    BaixaNotaAbastecimento.ValorPago = fValidaValor2(txt_valor_baixado.Text)
    lEmpresa_baixa = BaixaNotaAbastecimento.Empresa
    lDataAbastecimento = BaixaNotaAbastecimento.DataAbastecimento
    lDataPagamento = BaixaNotaAbastecimento.DataPagamento
    lPeriodo = BaixaNotaAbastecimento.Periodo
    lNumeroNota = BaixaNotaAbastecimento.NumeroNota
    lOrdem = BaixaNotaAbastecimento.Ordem
    lCodigoCliente = BaixaNotaAbastecimento.CodigoCliente
    lCodigoProduto = BaixaNotaAbastecimento.CodigoProduto2
    lTipoMovimento = BaixaNotaAbastecimento.TipoMovimento
End Sub
Private Sub AtualizaTelaNota()
    lbl_cliente.Caption = dtcboCliente.Text
    lbl_data_abastecimento.Caption = Format(MovNotaAbastecimento.DataAbastecimento, "dd/mm/yyyy")
    lbl_periodo.Caption = MovNotaAbastecimento.Periodo
    If MovNotaAbastecimento.TipoMovimento = 1 Then
        lbl_tipo_movimento.Caption = "1 - Caixa de combustíveis"
    Else
        lbl_tipo_movimento.Caption = "2 - Caixa de óleo/diversos"
    End If
    lbl_numero_nota.Caption = MovNotaAbastecimento.NumeroNota
    lblOrdem.Caption = MovNotaAbastecimento.Ordem
    If Produto.LocalizarCodigo(MovNotaAbastecimento.CodigoProduto2) Then
        lbl_codigo_produto.Caption = Produto.Codigo
        lbl_nome_produto.Caption = Produto.Nome
    End If
    lbl_quantidade.Caption = Format(MovNotaAbastecimento.Quantidade, "###,##0.00")
    lbl_total.Caption = Format(MovNotaAbastecimento.ValorTotal, "###,##0.00")
    msk_data_baixa.Text = Format(g_data, "dd/mm/yyyy")
    txt_valor_baixado.Text = Format(MovNotaAbastecimento.ValorTotal, "###,##0.00")
    lOpcao = 1
End Sub
Private Sub AtualTela()
    lEmpresa_baixa = BaixaNotaAbastecimento.Empresa
    lDataAbastecimento = BaixaNotaAbastecimento.DataAbastecimento
    lPeriodo = BaixaNotaAbastecimento.Periodo
    lNumeroNota = BaixaNotaAbastecimento.NumeroNota
    lOrdem = BaixaNotaAbastecimento.Ordem
    lCodigoCliente = BaixaNotaAbastecimento.CodigoCliente
    lCodigoProduto = BaixaNotaAbastecimento.CodigoProduto2
    lTipoMovimento = BaixaNotaAbastecimento.TipoMovimento
    lDataPagamento = BaixaNotaAbastecimento.DataPagamento
    lNumeroMovimentoCaixa = BaixaNotaAbastecimento.NumeroMovimentoCaixaBaixa
    
    If Cliente.LocalizarCodigo(BaixaNotaAbastecimento.CodigoCliente) Then
        lbl_cliente.Caption = Cliente.RazaoSocial
    Else
        lbl_cliente.Caption = "** Cliente Não Cadastrado **"
    End If
    lbl_data_abastecimento.Caption = Format(BaixaNotaAbastecimento.DataAbastecimento, "dd/mm/yyyy")
    lbl_periodo.Caption = BaixaNotaAbastecimento.Periodo
    If BaixaNotaAbastecimento.TipoMovimento = 1 Then
        lbl_tipo_movimento.Caption = "1 - Caixa de combustíveis"
    Else
        lbl_tipo_movimento.Caption = "2 - Caixa de óleo/diversos"
    End If
    lbl_numero_nota.Caption = BaixaNotaAbastecimento.NumeroNota
    lblOrdem.Caption = BaixaNotaAbastecimento.Ordem
    If Produto.LocalizarCodigo(BaixaNotaAbastecimento.CodigoProduto2) Then
        lbl_codigo_produto.Caption = Produto.Codigo
        lbl_nome_produto.Caption = Produto.Nome
    End If
    lbl_quantidade.Caption = Format(BaixaNotaAbastecimento.Quantidade, "###,##0.00")
    lbl_total.Caption = Format(BaixaNotaAbastecimento.ValorTotal, "###,##0.00")
    msk_data_baixa.Text = Format(BaixaNotaAbastecimento.DataPagamento, "dd/mm/yyyy")
    txt_valor_baixado.Text = Format(BaixaNotaAbastecimento.ValorPago, "###,##0.00")
    frmDados.Enabled = False
End Sub
Function BuscaBaixaCliente() As Boolean
    BuscaBaixaCliente = False
    If BaixaNotaAbastecimento.LocalizarUltimo(Val(txt_cliente.Text)) Then
        BuscaBaixaCliente = True
        cmd_alterar.Enabled = True
        cmd_extornar.Enabled = True
        frm_move.Enabled = True
        AtualTela
    Else
        cmd_alterar.Enabled = False
        cmd_extornar.Enabled = False
        frm_move.Enabled = False
        LimpaTela
    End If
End Function
Private Sub DesativaBotoes()
    cmd_alterar.Enabled = False
    cmd_extornar.Enabled = False
    cmd_pesquisa.Enabled = False
    cmd_sair.Enabled = False
    frm_move.Visible = False
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
End Sub
Private Sub ExcluiMovimentoCaixa()
    If Not MovCaixa.Excluir(g_empresa, lDataPagamento, lNumeroMovimentoCaixa) Then
        MsgBox "Não foi excluído o movimento do caixa!", vbInformation, "Erro de Integridade."
    End If
End Sub
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set BaixaNotaAbastecimento = Nothing
    Set Cliente = Nothing
    Set ConfiguracaoDiversa = Nothing
    Set IntegracaoCaixa = Nothing
    Set MovCaixa = Nothing
    Set MovNotaAbastecimento = Nothing
    Set Produto = Nothing
End Sub
Private Sub cmd_alterar_Click()
    Call GravaAuditoria(1, Me.name, 3, "")
    lOpcao = 2
    If BaixaNotaAbastecimento.LocalizarCodigo(g_empresa, lCodigoCliente, lDataAbastecimento, lNumeroNota, lOrdem, lCodigoProduto, lPeriodo) Then
        DesativaBotoes
        cmd_ok.Visible = True
        cmd_cancelar.Visible = True
        frmDados.Enabled = True
        msk_data_baixa.SetFocus
    Else
        MsgBox "Erro de verificação."
    End If
End Sub
Private Sub cmd_anterior_Click()
    Call GravaAuditoria(1, Me.name, 13, "")
    If BaixaNotaAbastecimento.LocalizarAnterior(Val(txt_cliente.Text)) Then
        AtualTela
    Else
        MsgBox "Início de Arquivo.", vbExclamation, "Atenção!"
        cmd_proximo.SetFocus
    End If
End Sub
Private Sub cmd_cancelar_Click()
    Call GravaAuditoria(1, Me.name, 9, "")
    If BuscaBaixaCliente Then
        AtivaBotoes
        'Atualtela
        cmd_alterar.SetFocus
        MSFlexGrid.SetFocus
    Else
        DesativaBotoes
        cmd_sair.Enabled = True
        'LimpaTela
        MSFlexGrid.SetFocus
    End If
End Sub
Private Sub LimpaTela()
    lbl_cliente.Caption = ""
    lbl_data_abastecimento.Caption = ""
    lbl_periodo.Caption = ""
    lbl_tipo_movimento.Caption = ""
    lbl_numero_nota.Caption = ""
    lblOrdem.Caption = ""
    lbl_codigo_produto.Caption = ""
    lbl_nome_produto.Caption = ""
    lbl_quantidade.Caption = ""
    lbl_total.Caption = ""
    msk_data_baixa.Text = "__/__/____"
    txt_valor_baixado.Text = ""
End Sub
Private Sub LimpaMSFlexGrid()
    Dim i As Integer
    MSFlexGrid.WordWrap = True
    MSFlexGrid.Rows = 2
    MSFlexGrid.Row = 1
    For i = 0 To 7
        MSFlexGrid.Col = i
        MSFlexGrid.Text = ""
    Next
    MSFlexGrid.RowHeight(0) = 650
    MSFlexGrid.Row = 0
    i = 0
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Data do Abastecimento"
    MSFlexGrid.ColWidth(i) = 1200
    MSFlexGrid.ColAlignment(i) = 4
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Per."
    MSFlexGrid.ColWidth(i) = 400
    MSFlexGrid.ColAlignment(i) = 4
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Tipo Mov."
    MSFlexGrid.ColWidth(i) = 500
    MSFlexGrid.ColAlignment(i) = 4
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Número da Nota"
    MSFlexGrid.ColWidth(i) = 800
    MSFlexGrid.ColAlignment(i) = 9
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Ordem"
    MSFlexGrid.ColWidth(i) = 500
    MSFlexGrid.ColAlignment(i) = 9
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Produto"
    MSFlexGrid.ColWidth(i) = 2500
    MSFlexGrid.ColAlignment(i) = 1
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Valor Total"
    MSFlexGrid.ColWidth(i) = 1100
    MSFlexGrid.ColAlignment(i) = 7
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Codigo do Produto"
    MSFlexGrid.ColWidth(i) = 700
    MSFlexGrid.ColAlignment(i) = 4
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Empresa"
    MSFlexGrid.ColWidth(i) = 700
    MSFlexGrid.ColAlignment(i) = 4
    MSFlexGrid.Row = 1
    MSFlexGrid.Col = 0
End Sub
Private Sub cmd_extornar_Click()
    Dim xString As String
    
    Call GravaAuditoria(1, Me.name, 19, "")
    If BaixaNotaAbastecimento.LocalizarCodigo(g_empresa, Val(txt_cliente.Text), CDate(lbl_data_abastecimento.Caption), Val(lbl_numero_nota.Caption), Val(lblOrdem.Caption), Val(lbl_codigo_produto.Caption), lbl_periodo.Caption) Then
        If (MsgBox("Deseja realmente extornar esta baixa?", 4 + 32 + 256, "Exclusão de Registro!")) = vbYes Then
            xString = "Cli:" & CLng(txt_cliente.Text)
            xString = xString & " Data:" & lbl_data_abastecimento.Caption & " Nota:" & lbl_numero_nota.Caption
            xString = xString & " Vlr:" & lbl_total.Caption
            Call GravaAuditoria(1, Me.name, 10, xString)
            xString = "Data Bx:" & msk_data_baixa.Text
            xString = xString & " Vlr Bx:" & txt_valor_baixado.Text
            Call GravaAuditoria(1, Me.name, 10, xString)
            If BaixaNotaAbastecimento.NumeroMovimentoCaixaBaixa > 0 Then
                Call ExcluiMovimentoCaixa
            End If
            AtualizaTabelaNota
            If MovNotaAbastecimento.Incluir Then
                If Not BaixaNotaAbastecimento.Excluir(g_empresa, Val(txt_cliente.Text), CDate(lbl_data_abastecimento.Caption), Val(lbl_numero_nota.Caption), Val(lblOrdem.Caption), Val(lbl_codigo_produto.Caption), lbl_periodo.Caption) Then
                    MsgBox "Não foi possível excluir este registro!", vbInformation, "Erro de Integridade"
                End If
            Else
                MsgBox "Não foi possível incluir este registro!", vbInformation, "Erro de Integridade"
            End If
            dtcboCliente_LostFocus
        End If
    End If
End Sub
Private Sub cmd_ok_Click()
    Dim xString As String
    
    On Error GoTo FileError
    
    If ValidaCampos Then
        AtivaBotoes
        If lOpcao = 1 Then
            xString = "Cli:" & CLng(txt_cliente.Text)
            xString = xString & " Data:" & lbl_data_abastecimento.Caption & " Nota:" & lbl_numero_nota.Caption
            xString = xString & " Vlr:" & lbl_total.Caption
            Call GravaAuditoria(1, Me.name, 10, xString)
            xString = "Data Bx:" & msk_data_baixa.Text
            xString = xString & " Vlr Bx:" & txt_valor_baixado.Text
            Call GravaAuditoria(1, Me.name, 10, xString)
            If MovNotaAbastecimento.BaixadoPelaDuplicata = False Then
                If Not IncluiMovimentoCaixa Then
                    MsgBox "Não foi possível integrar com o Caixa!", vbInformation, "Erro de Integridade."
                End If
            Else
                lNumeroMovimentoCaixa = 0
            End If
            AtualTabe
            g_data = BaixaNotaAbastecimento.DataAbastecimento
            If BaixaNotaAbastecimento.Incluir Then
                If Not MovNotaAbastecimento.Excluir(g_empresa, lCodigoCliente, lDataAbastecimento, lNumeroNota, lOrdem, lCodigoProduto, lPeriodo) Then
                    MsgBox "Não foi possível excluir este registro!", vbInformation, "Erro de Integridade"
                End If
            Else
                MsgBox "Não foi possível incluir este registro!", vbInformation, "Erro de Integridade"
            End If
        ElseIf lOpcao = 2 Then
            xString = "De Cli:" & CLng(txt_cliente.Text)
            xString = xString & " Data:" & lbl_data_abastecimento.Caption & " Nota:" & lbl_numero_nota.Caption
            xString = xString & " Vlr:" & lbl_total.Caption
            Call GravaAuditoria(1, Me.name, 10, xString)
            xString = "Data Bx:" & Format(BaixaNotaAbastecimento.DataPagamento, "dd/mm/yyyy")
            xString = xString & " Vlr Bx:" & Format(BaixaNotaAbastecimento.ValorPago, "###,##0.00")
            Call GravaAuditoria(1, Me.name, 10, xString)
            Atualtabe2
            xString = "Para Cli:" & CLng(txt_cliente.Text)
            xString = xString & " Data:" & lbl_data_abastecimento.Caption & " Nota:" & lbl_numero_nota.Caption
            xString = xString & " Vlr:" & lbl_total.Caption
            Call GravaAuditoria(1, Me.name, 10, xString)
            xString = "Data Bx:" & msk_data_baixa.Text
            xString = xString & " Vlr Bx:" & txt_valor_baixado.Text
            Call GravaAuditoria(1, Me.name, 10, xString)
            If Not BaixaNotaAbastecimento.Alterar(g_empresa, Val(txt_cliente.Text), CDate(lbl_data_abastecimento.Caption), Val(lbl_numero_nota.Caption), Val(lblOrdem.Caption), Val(lbl_codigo_produto.Caption), lbl_periodo.Caption) Then
                MsgBox "Não foi possível alterar este registro!", vbInformation, "Erro de Integridade"
            End If
        End If
        'AtualizaMSFlexGrid
        dtcboCliente_LostFocus
        If BaixaNotaAbastecimento.LocalizarCodigo(g_empresa, lCodigoCliente, lDataAbastecimento, lNumeroNota, lOrdem, lCodigoProduto, lPeriodo) Then
            AtualTela
        Else
            MsgBox "Não foi possível localizar o registro!", vbInformation, "Erro de Integridade"
        End If
        'MSFlexGrid.SetFocus
    End If
    Exit Sub
FileError:
    'ErroArquivo tbl_movimento_nota.Name, "Notaa"
    Exit Sub
End Sub
Function ValidaCampos() As Boolean
    ValidaCampos = False
    If Not IsDate(msk_data_baixa.Text) Then
        MsgBox "Informe a data da baixa.", vbInformation, "Atenção!"
        msk_data_baixa.SetFocus
    ElseIf CDate(msk_data_baixa.Text) < CDate(lbl_data_abastecimento.Caption) Then
        MsgBox "A data da baixa deve ser maior que " & lbl_data_abastecimento.Caption & ".", vbInformation, "Atenção!"
        msk_data_baixa.SetFocus
    ElseIf Not fValidaValor2(txt_valor_baixado.Text) > 0 Then
        MsgBox "Informe o valor baixado.", vbInformation, "Atenção!"
        txt_valor_baixado.SetFocus
    ElseIf fValidaValor2(txt_valor_baixado.Text) < fValidaValor2(lbl_total.Caption) Then
        MsgBox "O valor baixado não pode ser menor que " & lbl_total.Caption & ".", vbInformation, "Atenção!"
        txt_valor_baixado.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_pesquisa_Click()
    Call GravaAuditoria(1, Me.name, 5, "")
    g_valor = 1
    consulta_baixa_nota.Show 1
    If Len(g_string) > 0 Then
        lCodigoCliente = RetiraGString(1)
        lDataPagamento = RetiraGString(2)
        lCodigoProduto = RetiraGString(3)
        lNumeroNota = RetiraGString(4)
        lOrdem = RetiraGString(5)
        lDataAbastecimento = RetiraGString(6)
        lEmpresa = RetiraGString(7)
        lPeriodo = RetiraGString(8)
        If g_empresa = RetiraGString(7) Then
            If BaixaNotaAbastecimento.LocalizarCodigo(g_empresa, lCodigoCliente, lDataAbastecimento, lNumeroNota, lOrdem, lCodigoProduto, lPeriodo) Then
                cmd_alterar.Enabled = True
                cmd_extornar.Enabled = True
                txt_cliente.Text = lCodigoCliente
                dtcboCliente.BoundText = lCodigoCliente
                AtualTela
            Else
                MsgBox "Não foi possível localizar o registro.", vbInformation, "Erro de Integridade"
            End If
        Else
            MsgBox "A baixa selecionada é da empresa " & RetiraGString(7) & ".", vbInformation, "Atenção!"
        End If
    End If
End Sub
Private Sub cmd_primeiro_Click()
    Call GravaAuditoria(1, Me.name, 12, "")
    If BaixaNotaAbastecimento.LocalizarPrimeiro(Val(txt_cliente.Text)) Then
        AtualTela
        cmd_proximo.SetFocus
    Else
        MsgBox "Cliente não tem baixa.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    Call GravaAuditoria(1, Me.name, 14, "")
    If BaixaNotaAbastecimento.LocalizarProximo(Val(txt_cliente.Text)) Then
        AtualTela
    Else
        MsgBox "Fim de Arquivo.", vbExclamation, "Atenção!"
        cmd_anterior.SetFocus
    End If
End Sub
Private Sub cmd_sair_Click()
    Finaliza
    Unload Me
End Sub
Private Sub cmd_ultimo_Click()
    Call GravaAuditoria(1, Me.name, 15, "")
    If BaixaNotaAbastecimento.LocalizarUltimo(Val(txt_cliente.Text)) Then
        AtualTela
        cmd_anterior.SetFocus
    Else
        MsgBox "Cliente não tem baixa.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub dtcboCliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        MSFlexGrid.SetFocus
    End If
End Sub
Private Sub dtcboCliente_LostFocus()
    If dtcboCliente.BoundText <> "" Then
        If Cliente.LocalizarCodigo(Val(dtcboCliente.BoundText)) Then
            If Cliente.Codigo <> Val(txt_cliente.Text) Then
                txt_cliente.Text = Cliente.Codigo
            End If
            BuscaBaixaCliente
            AtualizaMSFlexGrid
            MSFlexGrid.SetFocus
        End If
    End If
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    If flag_baixa_nota_abastecimento = 0 Then
        lOpcao = 0
        DesativaBotoes
        If BaixaNotaAbastecimento.LocalizarUltimoRegistro() Then
            AtualTela
            AtivaBotoes
            AtualizaMSFlexGrid
        Else
            cmd_sair.Enabled = True
        End If
        txt_cliente.SetFocus
    Else
        flag_baixa_nota_abastecimento = 0
    End If
    Screen.MousePointer = 1
    If ConfiguracaoDiversa.LocalizarCodigo(1, "Baixar Notas Abast. Pelo Financeiro") Then
        If ConfiguracaoDiversa.Verdadeiro Then
            MsgBox "Esta operação somente será aceita pelo movimento financeiro.", vbInformation + vbOKOnly + vbExclamation, "Operação não Permitida!"
            cmd_sair_Click
        End If
    Else
        MsgBox "Esta operação somente será aceita pelo movimento financeiro.", vbInformation + vbOKOnly + vbExclamation, "Operação não Permitida!"
        cmd_sair_Click
    End If
End Sub
Private Sub MarcaCelulas()
    Call GravaAuditoria(1, Me.name, 18, "")
    If MSFlexGrid.TextMatrix(MSFlexGrid.Row, 0) <> "" Then
        lCodigoCliente = Val(txt_cliente.Text)
        lDataAbastecimento = MSFlexGrid.TextMatrix(MSFlexGrid.Row, 0)
        lPeriodo = MSFlexGrid.TextMatrix(MSFlexGrid.Row, 1)
        lTipoMovimento = MSFlexGrid.TextMatrix(MSFlexGrid.Row, 2)
        lNumeroNota = MSFlexGrid.TextMatrix(MSFlexGrid.Row, 3)
        lOrdem = MSFlexGrid.TextMatrix(MSFlexGrid.Row, 4)
        lCodigoProduto = MSFlexGrid.TextMatrix(MSFlexGrid.Row, 7)
        lEmpresa = MSFlexGrid.TextMatrix(MSFlexGrid.Row, 8)
        Inclui
        If MovNotaAbastecimento.LocalizarCodigo(g_empresa, lCodigoCliente, lDataAbastecimento, lNumeroNota, lOrdem, lCodigoProduto, lPeriodo) Then
            AtualizaTelaNota
            cmd_ok.SetFocus
        End If
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 And lOpcao = 0 Then
        KeyCode = 0
        cmd_alterar_Click
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
    
    Set adodcCliente.Recordset = Conectar.RsConexao("SELECT Codigo, [Razao Social] FROM Cliente WHERE Inativo = " & preparaBooleano(False) & " ORDER BY [Razao Social]")
    'dta_cliente.RecordSource = "SELECT Codigo, [Razao Social] FROM Cliente WHERE Inativo = " & preparaBooleano(False) & " ORDER BY [Razao Social]"
    'dta_cliente.Refresh
    g_data = g_data_def
End Sub
Private Sub MSFlexGrid_DblClick()
    MarcaCelulas
End Sub
Private Sub MSFlexGrid_KeyDown(KeyCode As Integer, Shift As Integer)
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
Private Sub MSFlexGrid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
        KeyAscii = 0
        MarcaCelulas
    End If
End Sub

Private Sub msk_data_baixa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_valor_baixado.SetFocus
    End If
End Sub
Private Sub txt_cliente_GotFocus()
    txt_cliente.SelStart = 0
    txt_cliente.SelLength = Len(txt_cliente.Text)
End Sub
Private Sub txt_cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dtcboCliente.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_cliente_LostFocus()
    If Val(txt_cliente.Text) > 0 Then
        If Cliente.LocalizarCodigo(Val(txt_cliente.Text)) Then
            dtcboCliente.BoundText = Val(txt_cliente.Text)
            MSFlexGrid.SetFocus
            Exit Sub
        Else
            MsgBox "Cliente não cadastro.", vbInformation, "Atenção!"
            txt_cliente.SetFocus
            Exit Sub
        End If
    End If
End Sub
Private Sub txt_valor_baixado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    End If
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub txt_valor_baixado_LostFocus()
    txt_valor_baixado.Text = Format(txt_valor_baixado.Text, "###,##0.00")
End Sub
