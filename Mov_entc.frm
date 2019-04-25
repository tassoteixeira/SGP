VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form mov_entrada_combustiveis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entrada de Combustíveis"
   ClientHeight    =   7185
   ClientLeft      =   4785
   ClientTop       =   1470
   ClientWidth     =   8370
   Icon            =   "Mov_entc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "Mov_entc.frx":030A
   ScaleHeight     =   7185
   ScaleWidth      =   8370
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3720
      Picture         =   "Mov_entc.frx":0750
      Style           =   1  'Graphical
      TabIndex        =   54
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   6240
      Width           =   795
   End
   Begin VB.CommandButton cmd_pesquisa 
      Caption         =   "&Pesquisa"
      Height          =   855
      Left            =   2820
      Picture         =   "Mov_entc.frx":1DE2
      Style           =   1  'Graphical
      TabIndex        =   53
      ToolTipText     =   "Pesquisa um registro específico."
      Top             =   6240
      Width           =   795
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   1920
      Picture         =   "Mov_entc.frx":3254
      Style           =   1  'Graphical
      TabIndex        =   52
      ToolTipText     =   "Exclui o registro atual."
      Top             =   6240
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   1020
      Picture         =   "Mov_entc.frx":48E6
      Style           =   1  'Graphical
      TabIndex        =   51
      ToolTipText     =   "Altera o registro atual."
      Top             =   6240
      Width           =   795
   End
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   120
      Picture         =   "Mov_entc.frx":5DE0
      Style           =   1  'Graphical
      TabIndex        =   50
      ToolTipText     =   "Cria um novo registro."
      Top             =   6240
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Enabled         =   0   'False
      Height          =   6075
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   8175
      Begin VB.TextBox txtDataEmissao 
         Height          =   285
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.ComboBox cboCst 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   4920
         Width           =   3555
      End
      Begin VB.TextBox txtValorIcmsSt 
         Height          =   300
         Left            =   6960
         MaxLength       =   12
         TabIndex        =   42
         Top             =   5280
         Width           =   1095
      End
      Begin VB.TextBox txtValorBcIcmsSt 
         Height          =   300
         Left            =   2400
         MaxLength       =   12
         TabIndex        =   40
         Top             =   5280
         Width           =   1095
      End
      Begin VB.TextBox txtValorIcms 
         Height          =   300
         Left            =   2400
         MaxLength       =   12
         TabIndex        =   36
         Top             =   4560
         Width           =   1095
      End
      Begin VB.TextBox txtAliquotaIcms 
         Height          =   300
         Left            =   6960
         MaxLength       =   12
         TabIndex        =   34
         Top             =   4200
         Width           =   1095
      End
      Begin VB.TextBox txtValorBcIcms 
         Height          =   300
         Left            =   2400
         MaxLength       =   12
         TabIndex        =   32
         Top             =   4200
         Width           =   1095
      End
      Begin VB.ComboBox cboFormaPagamento 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1320
         Width           =   2355
      End
      Begin VB.TextBox txtValorFrete 
         Height          =   300
         Left            =   6960
         MaxLength       =   12
         TabIndex        =   46
         Top             =   5640
         Width           =   1095
      End
      Begin VB.ComboBox cboTanque 
         Height          =   315
         Left            =   6960
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   3120
         Width           =   975
      End
      Begin VB.TextBox txtChaveAcesso 
         Height          =   285
         Left            =   2400
         MaxLength       =   44
         TabIndex        =   10
         Top             =   960
         Width           =   5655
      End
      Begin VB.TextBox txtSerie 
         Height          =   285
         Left            =   7500
         MaxLength       =   3
         TabIndex        =   8
         Top             =   600
         Width           =   555
      End
      Begin VB.TextBox txtModelo 
         Height          =   285
         Left            =   2400
         MaxLength       =   2
         TabIndex        =   6
         Top             =   600
         Width           =   435
      End
      Begin VB.TextBox txtItem 
         Height          =   300
         Left            =   2400
         MaxLength       =   2
         TabIndex        =   22
         Top             =   3120
         Width           =   375
      End
      Begin VB.TextBox txtData 
         Height          =   285
         Left            =   6960
         MaxLength       =   10
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_transfere_dados_lmc 
         Caption         =   "&Transfere p/ LMC"
         Height          =   675
         Left            =   5400
         Picture         =   "Mov_entc.frx":7472
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Transfere as entradas de combustíveis para o LMC."
         Top             =   2040
         Width           =   1455
      End
      Begin VB.TextBox txtValorTotal 
         Height          =   300
         Left            =   2400
         MaxLength       =   12
         TabIndex        =   30
         Top             =   3840
         Width           =   1095
      End
      Begin VB.ComboBox cbo_transporte 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   2400
         Width           =   2355
      End
      Begin VB.TextBox txt_nota 
         Height          =   300
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   16
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox txtQuantidade 
         Height          =   300
         Left            =   6960
         MaxLength       =   12
         TabIndex        =   28
         Top             =   3480
         Width           =   1095
      End
      Begin VB.TextBox txtValorLitro 
         Height          =   300
         Left            =   2400
         MaxLength       =   12
         TabIndex        =   26
         Top             =   3480
         Width           =   1095
      End
      Begin VB.ComboBox cbo_combustivel 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   2760
         Width           =   5655
      End
      Begin MSAdodcLib.Adodc adodc_fornecedor 
         Height          =   330
         Left            =   4560
         Top             =   1680
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
      Begin MSDataListLib.DataCombo dtcbo_fornecedor 
         Bindings        =   "Mov_entc.frx":8864
         DataSource      =   "adodc_fornecedor"
         Height          =   315
         Left            =   2400
         TabIndex        =   14
         Top             =   1680
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   "dtcbo_fornecedor"
      End
      Begin VB.TextBox txtValorNaoTributadoRedBcIcms 
         Height          =   300
         Left            =   2400
         MaxLength       =   12
         TabIndex        =   44
         Top             =   5640
         Width           =   1095
      End
      Begin VB.Label Label22 
         Caption         =   "Data da &Emissão"
         Height          =   300
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2235
      End
      Begin VB.Label Label23 
         Caption         =   "CST"
         Height          =   300
         Left            =   120
         TabIndex        =   37
         Top             =   4920
         Width           =   2235
      End
      Begin VB.Label Label21 
         Caption         =   "Vl.Nao Tributado Red.Bc Icms"
         Height          =   300
         Left            =   120
         TabIndex        =   43
         Top             =   5640
         Width           =   2235
      End
      Begin VB.Label Label20 
         Caption         =   "Valor do ICMS ST"
         Height          =   300
         Left            =   5280
         TabIndex        =   41
         Top             =   5280
         Width           =   1635
      End
      Begin VB.Label Label19 
         Caption         =   "Valor Base Calculo ICMS ST"
         Height          =   300
         Left            =   120
         TabIndex        =   39
         Top             =   5280
         Width           =   2235
      End
      Begin VB.Label Label18 
         Caption         =   "Valor do ICMS"
         Height          =   300
         Left            =   120
         TabIndex        =   35
         Top             =   4560
         Width           =   2235
      End
      Begin VB.Label Label17 
         Caption         =   "Aliquota do ICMS"
         Height          =   300
         Left            =   5280
         TabIndex        =   33
         Top             =   4200
         Width           =   1635
      End
      Begin VB.Label Label16 
         Caption         =   "Valor da Base de Calculo ICMS"
         Height          =   300
         Left            =   120
         TabIndex        =   31
         Top             =   4200
         Width           =   2235
      End
      Begin VB.Label Label15 
         Caption         =   "Forma de Pagamento"
         Height          =   300
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   2235
      End
      Begin VB.Label Label14 
         Caption         =   "Valor do Frete"
         Height          =   300
         Left            =   5280
         TabIndex        =   45
         Top             =   5640
         Width           =   1635
      End
      Begin VB.Label Label13 
         Caption         =   "Número do Tanq&ue"
         Height          =   300
         Left            =   5280
         TabIndex        =   23
         Top             =   3120
         Width           =   1635
      End
      Begin VB.Label Label12 
         Caption         =   "C&have de Acesso"
         Height          =   300
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   2235
      End
      Begin VB.Label Label11 
         Caption         =   "Sé&rie da NF"
         Height          =   300
         Left            =   5940
         TabIndex        =   7
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "&Modelo da NF"
         Height          =   300
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   2235
      End
      Begin VB.Label Label7 
         Caption         =   "Ítem"
         Height          =   300
         Left            =   120
         TabIndex        =   21
         Top             =   3120
         Width           =   2235
      End
      Begin VB.Label Label3 
         Caption         =   "&Fornecedor"
         Height          =   300
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   2235
      End
      Begin VB.Label Label10 
         Caption         =   "&Tipo de Transporte"
         Height          =   300
         Left            =   120
         TabIndex        =   17
         Top             =   2400
         Width           =   2235
      End
      Begin VB.Label Label8 
         Caption         =   "Valor Total"
         Height          =   300
         Left            =   120
         TabIndex        =   29
         Top             =   3840
         Width           =   2235
      End
      Begin VB.Label Label6 
         Caption         =   "&Quantidade de Litros"
         Height          =   300
         Left            =   5280
         TabIndex        =   27
         Top             =   3480
         Width           =   1635
      End
      Begin VB.Label Label1 
         Caption         =   "&Valor de Custo"
         Height          =   300
         Left            =   120
         TabIndex        =   25
         Top             =   3480
         Width           =   2235
      End
      Begin VB.Label Label4 
         Caption         =   "&Número da N.F."
         Height          =   300
         Left            =   120
         TabIndex        =   15
         Top             =   2040
         Width           =   2235
      End
      Begin VB.Label Label5 
         Caption         =   "&Data da Entrada"
         Height          =   300
         Left            =   4680
         TabIndex        =   3
         Top             =   240
         Width           =   2235
      End
      Begin VB.Label Label2 
         Caption         =   "Co&mbustível"
         Height          =   300
         Left            =   120
         TabIndex        =   19
         Top             =   2760
         Width           =   2235
      End
   End
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   6120
      TabIndex        =   55
      Top             =   6120
      Width           =   2175
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "Mov_entc.frx":8883
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "Mov_entc.frx":9E05
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "Mov_entc.frx":B277
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "Mov_entc.frx":C771
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   6600
      Picture         =   "Mov_entc.frx":DC6B
      Style           =   1  'Graphical
      TabIndex        =   48
      ToolTipText     =   "Confirma o registro atual."
      Top             =   6240
      Width           =   795
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   7500
      Picture         =   "Mov_entc.frx":F275
      Style           =   1  'Graphical
      TabIndex        =   49
      ToolTipText     =   "Cancela o registro atual."
      Top             =   6240
      Width           =   795
   End
End
Attribute VB_Name = "mov_entrada_combustiveis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flag_movimento_Entrada_combustivel As Integer
Dim lOpcao As Integer
Dim lTipoCombustivel As String
Dim lData As Date
Dim lNota As String
Dim lCodigoFornecedor As Integer
Dim lItem As Integer
Dim lOrdem As Integer
Dim lTipoCombustivelAnt As String
Dim lQuantidadeAnt As Currency
Dim lCodigoFornecedorAnt As Integer
Dim lTotalAnt As Currency
Dim lChaveAcesso As String
Dim lCstAlcool060 As Boolean

Private rsTabela As New adodb.Recordset

Private Bomba As New cBomba
Private Combustivel As New cCombustivel
Private ConfiguracaoDiversa As New cConfiguracaoDiversa
Private EntradaCombustivel As New cEntradaCombustivel
Private Fornecedor As New cFornecedor
Private LivroLMC As New cLivroLMC
Private Produto As New cProduto
Private TanqueCombustivel As New cTanqueCombustivel
Private IntegracaoNuvem As New cIntegracaoNuvem
Private UnidadeFederacao As New cUnidadeFederacao

Private Sub AdicionaEstoque(x_tipo_combustivel As String, x_quantidade As Currency)
    If Combustivel.LocalizarCodigo(g_empresa, x_tipo_combustivel) Then
        Combustivel.QuantidadeEmEstoque = Combustivel.QuantidadeEmEstoque + x_quantidade
        If Not Combustivel.Alterar(g_empresa, x_tipo_combustivel) Then
            MsgBox "Não foi possível adicionar ao estoque", vbInformation, "Registro Não Encontrado"
        End If
    End If
End Sub
Private Sub AtivaBotoes()
    cmd_novo.Enabled = True
    cmd_excluir.Enabled = True
    cmd_alterar.Enabled = True
    cmd_pesquisa.Enabled = True
    cmd_sair.Enabled = True
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
    cmd_transfere_dados_lmc.Visible = False
    frm_move.Visible = True
End Sub
Private Sub AtualizaPrecoBicos()
    Dim xSQL As String
    Dim xResposta As Integer
    
    xResposta = 0
    'Prepara SQL
    xSQL = "SELECT Codigo"
    xSQL = xSQL & "  FROM Bomba"
    xSQL = xSQL & " WHERE Empresa = " & g_empresa
    xSQL = xSQL & "   AND [Tipo de Combustivel] = " & preparaTexto(lTipoCombustivel)
    xSQL = xSQL & " ORDER BY Codigo"
    'Abre RecordSet
    Set rsTabela = New adodb.Recordset
    Set rsTabela = Conectar.RsConexao(xSQL)
    'Verifica tabela
    If rsTabela.RecordCount > 0 Then
        Do Until rsTabela.EOF
            If Bomba.LocalizarCodigo(g_empresa, rsTabela("Codigo").Value) Then
                If Bomba.PrecoCusto <> fValidaValor(txtValorLitro.Text) Then
                    If xResposta = 0 Then
                        xResposta = (MsgBox("O preço de custo informado nesta nota é: " & txtValorLitro.Text & Chr(10) & Chr(13) & Chr(10) & "E o preço de custo do bico " & Bomba.Codigo & " é: " & Format(Bomba.PrecoCusto, "###,###,##0.0000") & Chr(10) & Chr(13) & Chr(10) & "Deseja alterar o preço de custo nos cadastros de Bomba/Bico e Produtos?", vbYesNo + vbDefaultButton1, "Alteração de Preço de Custo"))
                    End If
                    If xResposta = vbYes Then
                        Bomba.PrecoCusto = fValidaValor(txtValorLitro.Text)
                        If Not Bomba.Alterar(g_empresa, rsTabela("Codigo").Value) Then
                            MsgBox "Não foi possível alterar o bico:" & rsTabela("Codigo").Value, vbInformation, "Erro de Integridade!"
                        End If
                        If Produto.LocalizarCodigo(Bomba.CodigoProduto) Then
                            Produto.PrecoCusto = fValidaValor(txtValorLitro.Text)
                            If Not Produto.AlterarCusto(Bomba.CodigoProduto, g_empresa, Produto.PrecoCusto) Then
                                MsgBox "Não foi possível alterar o produto:" & Bomba.CodigoProduto, vbInformation, "Erro de Integridade!"
                            End If
                        Else
                            MsgBox "Não existe produto com o código:" & Bomba.CodigoProduto, vbInformation, "Erro de Integridade!"
                        End If
                    End If
                End If
            End If
            rsTabela.MoveNext
        Loop
    End If
    rsTabela.Close
    Set rsTabela = Nothing
End Sub
Private Sub AtualizaTabela()
    EntradaCombustivel.Empresa = g_empresa
    EntradaCombustivel.Data = CDate(txtData.Text)
    EntradaCombustivel.CodigoFornecedor = Val(dtcbo_fornecedor.BoundText)
    EntradaCombustivel.NumeroNota = Val(txt_nota.Text)
    EntradaCombustivel.TipoTransporte = Mid(cbo_transporte.Text, 1, 1)
    EntradaCombustivel.TipoCombustivel = Mid(cbo_combustivel.Text, 1, 2)
    EntradaCombustivel.ValorLitro = fValidaValor(txtValorLitro.Text)
    EntradaCombustivel.Quantidade = fValidaValor(txtQuantidade.Text)
    EntradaCombustivel.ValorEntrada = fValidaValor(txtValorTotal.Text)
    EntradaCombustivel.Item = Val(txtItem.Text)
    EntradaCombustivel.Modelo = txtModelo.Text
    EntradaCombustivel.Serie = txtSerie.Text
    EntradaCombustivel.ChaveAcesso = txtChaveAcesso.Text
    EntradaCombustivel.NumeroTanque = Val(cboTanque.Text)
    EntradaCombustivel.FormaPagamento = Mid(cboFormaPagamento.Text, 1, 1)
    EntradaCombustivel.ValorFrete = fValidaValor(txtValorFrete.Text)
    If lOpcao = 1 Then
        EntradaCombustivel.Ordem = 0 'Internamente a clase vai dar sequencia
    Else
        EntradaCombustivel.Ordem = lOrdem
    End If
    EntradaCombustivel.CST = Mid(cboCst.Text, 1, 3)
    EntradaCombustivel.Cfop = "1652"
    If Fornecedor.LocalizarCodigo(g_empresa, EntradaCombustivel.CodigoFornecedor) Then
        If UCase(gUfEmpresa) <> UCase(Fornecedor.UF) Then
            EntradaCombustivel.Cfop = "2652"
        End If
    End If
    EntradaCombustivel.CodigoProduto = 0
    If Bomba.LocalizarTipoCombustivel(g_empresa, EntradaCombustivel.TipoCombustivel) Then
        EntradaCombustivel.CodigoProduto = Bomba.CodigoProduto
    End If
    EntradaCombustivel.ValorBCICMS = fValidaValor(txtValorBcIcms.Text)
    EntradaCombustivel.AliquotaICMS = fValidaValor(txtAliquotaIcms.Text)
    EntradaCombustivel.ValorICMS = fValidaValor(txtValorIcms.Text)
    EntradaCombustivel.ValorBCICMSST = fValidaValor(txtValorBcIcmsSt.Text)
    EntradaCombustivel.ValorICMSST = fValidaValor(txtValorIcmsSt.Text)
    EntradaCombustivel.ValorNaoTributadoReducaoBCICMS = fValidaValor(txtValorNaoTributadoRedBcIcms.Text)
    EntradaCombustivel.DataEmissao = CDate(txtDataEmissao.Text)
End Sub
Private Sub AtualizaTela()
    Dim i As Integer
    
    lData = EntradaCombustivel.Data
    lTipoCombustivel = EntradaCombustivel.TipoCombustivel
    lNota = EntradaCombustivel.NumeroNota
    lCodigoFornecedor = EntradaCombustivel.CodigoFornecedor
    lItem = EntradaCombustivel.Item
    lOrdem = EntradaCombustivel.Ordem
    lTipoCombustivelAnt = EntradaCombustivel.TipoCombustivel
    lQuantidadeAnt = EntradaCombustivel.Quantidade
    lCodigoFornecedorAnt = EntradaCombustivel.CodigoFornecedor
    lTotalAnt = EntradaCombustivel.ValorEntrada
    lChaveAcesso = EntradaCombustivel.ChaveAcesso
    
    txtData.Text = Format(EntradaCombustivel.Data, "dd/mm/yyyy")
    txtDataEmissao.Text = Format(EntradaCombustivel.DataEmissao, "dd/mm/yyyy")
    dtcbo_fornecedor.BoundText = EntradaCombustivel.CodigoFornecedor
    For i = 0 To cbo_transporte.ListCount - 1
        cbo_transporte.ListIndex = i
        If Mid(cbo_transporte.Text, 1, 1) = EntradaCombustivel.TipoTransporte Then
            Exit For
        End If
    Next
    txtChaveAcesso.Text = EntradaCombustivel.ChaveAcesso
    For i = 0 To cboFormaPagamento.ListCount - 1
        cboFormaPagamento.ListIndex = i
        If Mid(cboFormaPagamento.Text, 1, 1) = EntradaCombustivel.FormaPagamento Then
            Exit For
        End If
    Next
    If Combustivel.LocalizarCodigo(g_empresa, EntradaCombustivel.TipoCombustivel) Then
        For i = 0 To cbo_combustivel.ListCount - 1
            cbo_combustivel.ListIndex = i
            If Mid(cbo_combustivel.Text, 1, 2) = Combustivel.Codigo Then
                Exit For
            End If
        Next
    Else
        cbo_combustivel.ListIndex = -1
    End If
    txtItem.Text = Format(EntradaCombustivel.Item, "00")
    txt_nota.Text = EntradaCombustivel.NumeroNota
    txtValorLitro.Text = Format(EntradaCombustivel.ValorLitro, "###,##0.0000")
    txtQuantidade.Text = Format(EntradaCombustivel.Quantidade, "###,##0.0;-###,##0.0")
    txtValorTotal.Text = Format(EntradaCombustivel.ValorEntrada, "###,###,##0.00")
    txtValorBcIcms.Text = Format(EntradaCombustivel.ValorBCICMS, "###,###,##0.00")
    txtAliquotaIcms.Text = Format(EntradaCombustivel.AliquotaICMS, "###,###,##0.00")
    txtValorIcms.Text = Format(EntradaCombustivel.ValorICMS, "###,###,##0.00")
    For i = 0 To cboCst.ListCount - 1
        cboCst.ListIndex = i
        If Mid(cboCst.Text, 1, 3) = EntradaCombustivel.CST Then
            Exit For
        End If
    Next
    txtValorBcIcmsSt.Text = Format(EntradaCombustivel.ValorBCICMSST, "###,###,##0.00")
    txtValorIcmsSt.Text = Format(EntradaCombustivel.ValorICMSST, "###,###,##0.00")
    txtValorNaoTributadoRedBcIcms.Text = Format(EntradaCombustivel.ValorNaoTributadoReducaoBCICMS, "###,###,##0.00")
    txtValorFrete.Text = Format(EntradaCombustivel.ValorFrete, "###,###,##0.00")
    If Not Bomba.LocalizarTipoCombustivel(g_empresa, EntradaCombustivel.TipoCombustivel) Then
        MsgBox "Combustível inexistente", vbInformation, "Erro de Consistência de Dados"
    End If
    txtModelo.Text = EntradaCombustivel.Modelo
    txtSerie.Text = EntradaCombustivel.Serie
    PreencheCboTanque
    For i = 0 To cboTanque.ListCount - 1
        cboTanque.ListIndex = i
        If Val(cboTanque.Text) = EntradaCombustivel.NumeroTanque Then
            Exit For
        End If
    Next
    If Val(cboTanque.Text) <> EntradaCombustivel.NumeroTanque Then
        MsgBox "O tanque informado não confere com o Mostrado na tela." & vbCrLf & "Tanque Informado=" & EntradaCombustivel.NumeroTanque & vbCrLf & "Favor alterar e confirmar esta Nota.", vbInformation, "Erro de Consistência de Dados"
    End If
    frm_dados.Enabled = False
    Call VerificaLiberacaoLMC(lTipoCombustivelAnt, lData)
End Sub
Private Function ChamaValidaChaveAcessoNFe() As Boolean
    Dim xResposta As String
    Dim xNumeroNF As Long
    
    ChamaValidaChaveAcessoNFe = False
    If CDate(txtData.Text) < CDate("01/07/2011") Then
        ChamaValidaChaveAcessoNFe = True
        Exit Function
    End If
    If Not Fornecedor.LocalizarCodigo(g_empresa, Val(dtcbo_fornecedor.BoundText)) Then
        MsgBox "Fornecedor não cadastrado.", vbInformation, "Erro de Integridade!"
        Exit Function
    End If
    xNumeroNF = fDesmascaraNumero(txt_nota.Text)
    xResposta = ValidaChaveAcessoNFe(txtChaveAcesso.Text, UCase(Fornecedor.UF), CDate(txtDataEmissao.Text), Fornecedor.CGC, txtSerie.Text, xNumeroNF)
    If xResposta = "OK" Then
        ChamaValidaChaveAcessoNFe = True
    Else
        MsgBox "Chave de Acesso Inválida." & vbCrLf & vbCrLf & "Erro:" & xResposta, vbInformation, "Atenção!"
        If txtModelo.Text <> "55" Then
            'Verifica Modelo do Documento
            If xResposta = "Modelo do documento deve ser 55" Then
                If MsgBox("Deseja que o sistema corrija o modelo do documento para o correto?", vbYesNo + vbQuestion + vbDefaultButton1, "Correção do Modelo do Documento!") = vbYes Then
                    txtModelo.Text = "55"
                Else
                    Exit Function
                End If
            Else
                txtModelo.Text = "55"
            End If
        End If
    End If
End Function
Private Sub DesativaBotoes()
    cmd_novo.Enabled = False
    cmd_excluir.Enabled = False
    cmd_alterar.Enabled = False
    cmd_pesquisa.Enabled = False
    cmd_sair.Enabled = False
    frm_move.Visible = False
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
    If g_nome_usuario = "L.M.C." Then
        cmd_transfere_dados_lmc.Visible = True
    End If
End Sub
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set Bomba = Nothing
    Set Combustivel = Nothing
    Set EntradaCombustivel = Nothing
    Set Fornecedor = Nothing
    Set LivroLMC = Nothing
    Set Produto = Nothing
    Set TanqueCombustivel = Nothing
    Set IntegracaoNuvem = Nothing
End Sub
Private Sub Inclui()
    lOpcao = 1
    DesativaBotoes
    cmd_novo.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    frm_dados.Enabled = True
    txtModelo.Text = "55"
    txtItem.Text = "01"
End Sub
Private Sub IncluiIntegracaoNuvem(ByVal pGeraRegistroExcluir As Boolean, ByVal pGeraIncluirCabecalho As Boolean)
    Dim xSQL As String
    
    If EntradaCombustivel.NomeTabela = "Entrada_Combustivel_LMC" Then
        IntegracaoNuvem.Empresa = g_empresa
        IntegracaoNuvem.Data = Format(Date, "dd/MM/yyyy") & " " & Format(Time, "HH:mm:SS")
        IntegracaoNuvem.NomeTabela = EntradaCombustivel.NomeTabela
        IntegracaoNuvem.ChaveAcesso = g_empresa & "|@|" & lData & "|@|" & "" & "|@|" & lNota & "|@|" & lCodigoFornecedor & "|@|" & 0 & "|@|"
        IntegracaoNuvem.TipoOperacao = " EXCLUIRCABECALHO"
        IntegracaoNuvem.IntegradoEm = CDate("00:00:00")
        If pGeraRegistroExcluir Then
            If Not IntegracaoNuvem.Incluir Then
                Call CriaLogSGP(Me.name & ".IncluiIntegracaoNuvem - Erro ao Incluir IntegracaoNuvem.", Err.Description, "Nota=" & lNota & " - Data=" & lData)
                Exit Sub
            End If
        End If

        If pGeraIncluirCabecalho Then
            'Verifica se Tem Registro da Nova Nota
            xSQL = "SELECT COUNT(1) AS Quantidade"
            xSQL = xSQL & "  FROM " & EntradaCombustivel.NomeTabela
            xSQL = xSQL & " WHERE Empresa = " & g_empresa
            xSQL = xSQL & "   AND Data = " & preparaData(txtData.Text)
            xSQL = xSQL & "   AND [Numero da Nota] = " & preparaTexto(txt_nota.Text)
            xSQL = xSQL & "   AND [Codigo do Fornecedor] = " & Val(dtcbo_fornecedor.BoundText)
            Set rsTabela = New adodb.Recordset
            Set rsTabela = Conectar.RsConexao(xSQL)
            If rsTabela.RecordCount > 0 Then
                If rsTabela("Quantidade").Value > 0 Then
                    IntegracaoNuvem.Data = Format(Date, "dd/MM/yyyy") & " " & Format(Time, "HH:mm:SS")
                    IntegracaoNuvem.ChaveAcesso = g_empresa & "|@|" & txtData.Text & "|@|" & "" & "|@|" & txt_nota.Text & "|@|" & Val(dtcbo_fornecedor.BoundText) & "|@|" & 0 & "|@|"
                    IntegracaoNuvem.TipoOperacao = "INCLUIRCABECALHO"
                    If Not IntegracaoNuvem.Incluir Then
                        Call CriaLogSGP(Me.name & ".IncluiIntegracaoNuvem - Erro ao Incluir IntegracaoNuvem..", Err.Description, "Nota=" & txt_nota.Text & " - Data=" & txtData.Text)
                    End If
                End If
            End If
            rsTabela.Close
        End If
        
        If pGeraRegistroExcluir Then
            If lData <> CDate(txtData.Text) Or lNota <> txt_nota.Text Or lTipoCombustivel <> Mid(cbo_combustivel.Text, 1, 2) Or lCodigoFornecedor <> Val(dtcbo_fornecedor.BoundText) Then
                'Verifica se Tem Registro da Nota Anterior
                xSQL = "SELECT COUNT(1) AS Quantidade"
                xSQL = xSQL & "  FROM " & EntradaCombustivel.NomeTabela
                xSQL = xSQL & " WHERE Empresa = " & g_empresa
                xSQL = xSQL & "   AND Data = " & preparaData(lData)
                xSQL = xSQL & "   AND [Numero da Nota] = " & preparaTexto(lNota)
                xSQL = xSQL & "   AND [Codigo do Fornecedor] = " & lCodigoFornecedor
                Set rsTabela = New adodb.Recordset
                Set rsTabela = Conectar.RsConexao(xSQL)
                If rsTabela.RecordCount > 0 Then
                    If rsTabela("Quantidade").Value > 0 Then
                        IntegracaoNuvem.Data = Format(Date, "dd/MM/yyyy") & " " & Format(Time, "HH:mm:SS")
                        IntegracaoNuvem.ChaveAcesso = g_empresa & "|@|" & lData & "|@|" & "" & "|@|" & lNota & "|@|" & lCodigoFornecedor & "|@|" & 0 & "|@|"
                        IntegracaoNuvem.TipoOperacao = "INCLUIR"
                        If Not IntegracaoNuvem.Incluir Then
                            Call CriaLogSGP(Me.name & ".IncluiIntegracaoNuvem - Erro ao Incluir IntegracaoNuvem..", Err.Description, "Nota=" & txt_nota.Text & " - Data=" & txtData.Text)
                        End If
                    End If
                End If
                rsTabela.Close
            End If
        End If
        Set rsTabela = Nothing
    End If
End Sub

Private Sub DesabilitaCamposTributacao()
    If Mid(cboCst.Text, 1, 3) = "060" Then
        txtValorBcIcmsSt.Text = "0,00"
        txtValorIcmsSt.Text = "0,00"
        txtValorNaoTributadoRedBcIcms.Text = "0,00"
        txtValorBcIcms.Text = "0,00"
        txtValorIcms.Text = "0,00"
        txtAliquotaIcms.Text = "0,00"
        
        txtValorBcIcmsSt.Enabled = False
        txtValorIcmsSt.Enabled = False
        txtValorNaoTributadoRedBcIcms.Enabled = False
        txtValorBcIcms.Enabled = False
        txtValorIcms.Enabled = False
        txtAliquotaIcms.Enabled = False
        txtValorFrete.SetFocus
    Else
        txtValorBcIcmsSt.Enabled = True
        txtValorIcmsSt.Enabled = True
        txtValorNaoTributadoRedBcIcms.Enabled = True
        txtValorBcIcms.Enabled = True
        txtValorIcms.Enabled = True
        txtAliquotaIcms.Enabled = True
        txtValorBcIcmsSt.SetFocus
    End If
End Sub
Private Sub cbo_combustivel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cboTanque.SetFocus
    End If
End Sub
Private Sub cbo_combustivel_LostFocus()
    Dim i As Integer
    Dim xCST As String
    
    If cbo_combustivel.ListIndex <> -1 Then
        If Not Bomba.LocalizarTipoCombustivel(g_empresa, Mid(cbo_combustivel.Text, 1, 2)) Then
            MsgBox "Bomba inexistente", vbInformation, "Erro de Consistência de Dados"
            cbo_combustivel.SetFocus
            Exit Sub
        End If
        If Not Combustivel.LocalizarCodigo(g_empresa, Mid(cbo_combustivel.Text, 1, 2)) Then
            MsgBox "Combustível inexistente", vbInformation, "Erro de Consistência de Dados"
            cbo_combustivel.SetFocus
            Exit Sub
        End If
        If lOpcao = 1 Then
            xCST = "060"
            If Mid(cbo_combustivel.Text, 1, 1) = "A" And lCstAlcool060 = False Then
                xCST = "070"
                If UCase(dtcbo_fornecedor.Text) Like "*ELLO*" Or VerificaFornecedorCst010 Then
                    xCST = "010"
                End If
            End If
            For i = 0 To cboCst.ListCount - 1
                cboCst.ListIndex = i
                If Mid(cboCst.Text, 1, 3) = xCST Then
                    Exit For
                End If
            Next
            PreencheCboTanque
            txtValorLitro.Text = Format(Bomba.PrecoCusto, "###0.0000")
            If IsDate(txtData.Text) And Mid(cbo_combustivel.Text, 1, 2) <> "" And Len(txt_nota.Text) > 0 And Val(txtItem.Text) > 0 Then
                If EntradaCombustivel.LocalizarCodigo(g_empresa, CDate(txtData.Text), Mid(cbo_combustivel.Text, 1, 2), txt_nota.Text, Val(txtItem.Text)) Then
                    If (MsgBox("Já existe uma nota com o produto selecionado." & vbCrLf & "Deseja lançar novamente o mesmo combustível na mesma nota?", vbYesNo + vbQuestion + vbDefaultButton2, "Duplicidade de Registro!")) = vbYes Then
                        txtItem.Text = Format(Val(txtItem.Text) + 1, "00")
                    Else
                        txtData.SetFocus
                    End If
                End If
            End If
        End If
    Else
        cbo_combustivel.SetFocus
    End If
End Sub
Private Sub cbo_transporte_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_combustivel.SetFocus
    End If
End Sub
Private Sub cboCst_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    DesabilitaCamposTributacao
'        If Mid(cboCst.Text, 1, 3) <> "060" Then
'            txtValorBcIcmsSt.Enabled = True
'            txtValorIcmsSt.Enabled = True
'            txtValorNaoTributadoRedBcIcms.Enabled = True
'            txtValorBcIcmsSt.SetFocus
'        Else
'            txtValorFrete.SetFocus
'        End If
    End If
End Sub

Private Sub cboCst_LostFocus()
 DesabilitaCamposTributacao
End Sub

Private Sub cboFormaPagamento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dtcbo_fornecedor.SetFocus
    End If
End Sub
Private Sub cboTanque_GotFocus()
    Dim xTanque As Integer
    Dim i As Integer
    
    xTanque = Val(cboTanque.Text)
    PreencheCboTanque
    For i = 0 To cboTanque.ListCount - 1
        cboTanque.ListIndex = i
        If Val(cboTanque.Text) = xTanque Then
            Exit For
        End If
    Next
End Sub
Private Sub cboTanque_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtValorLitro.SetFocus
    End If
End Sub
Private Sub cboTanque_LostFocus()
    If Val(cboTanque.Text) > 0 Then
        If Not TanqueCombustivel.LocalizarCodigo(g_empresa, Val(cboTanque.Text)) Then
            MsgBox "Tanque inexistente!", vbInformation, "Erro de Consistência de Dados"
            cboTanque.SetFocus
            Exit Sub
        End If
    Else
        cboTanque.SetFocus
    End If
End Sub
Private Sub cmd_alterar_Click()
    Call GravaAuditoria(1, Me.name, 3, "")
    lOpcao = 2
    DesativaBotoes
    cmd_alterar.Visible = True
    cmd_alterar.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    frm_dados.Enabled = True
    DesabilitaCamposTributacao
    txt_nota.SetFocus
    
'    If Mid(cboCst.Text, 1, 3) = "060" Then
'        txtValorBcIcmsSt.Enabled = False
'        txtValorIcmsSt.Enabled = False
'        txtValorNaoTributadoRedBcIcms.Enabled = False
'    End If
End Sub
Private Sub cmd_anterior_Click()
    Call GravaAuditoria(1, Me.name, 13, "")
    If EntradaCombustivel.LocalizarAnterior Then
        AtualizaTela
    Else
        MsgBox "Início de Arquivo.", 48, "Atenção!"
        cmd_proximo.SetFocus
    End If
End Sub
Private Sub cmd_cancelar_Click()
    Call GravaAuditoria(1, Me.name, 9, "")
    LimpaTela
    If EntradaCombustivel.LocalizarCodigo(g_empresa, lData, lTipoCombustivel, lNota, lOrdem) Then
        AtualizaTela
        AtivaBotoes
        cmd_alterar.SetFocus
    Else
        DesativaBotoes
        cmd_novo.Enabled = True
        cmd_sair.Enabled = True
        cmd_novo.SetFocus
    End If
    lOpcao = 0
End Sub
Private Sub LimpaTela()
    txtData.Text = ""
    txtModelo.Text = ""
    ''txtSerie.Text = ""
    txtChaveAcesso.Text = ""
    cboFormaPagamento.ListIndex = -1
    dtcbo_fornecedor.BoundText = ""
    txt_nota.Text = ""
    cbo_transporte.ListIndex = -1
    cbo_combustivel.ListIndex = -1
    txtItem.Text = ""
    txtValorLitro.Text = ""
    txtQuantidade.Text = ""
    txtValorTotal.Text = ""
    txtValorBcIcms.Text = "0,00"
    txtAliquotaIcms.Text = "0,00"
    txtValorIcms.Text = "0,00"
    cboCst.ListIndex = 1
    txtValorBcIcmsSt.Text = "0,00"
    txtValorIcmsSt.Text = "0,00"
    txtValorNaoTributadoRedBcIcms.Text = "0,00"
    txtValorFrete.Text = "0,00"
    txtDataEmissao.Text = ""
End Sub
Private Sub cmd_excluir_Click()
    Call GravaAuditoria(1, Me.name, 4, "")
    If IsDate(txtData.Text) Then
        If (MsgBox("Deseja Realmente Excluir Este Registro?", vbYesNo + vbQuestion + vbDefaultButton2, "Exclusão de Registro!")) = vbYes Then
            Call GravaAuditoria(1, Me.name, 10, "Data:" & txtData.Text & " Forn:" & Val(dtcbo_fornecedor.BoundText) & " NF:" & txt_nota.Text & " Comb:" & Mid(cbo_combustivel.Text, 1, 2) & " Qtd:" & txtQuantidade.Text & " Total:" & txtValorTotal.Text)
            If EntradaCombustivel.Excluir(g_empresa, CDate(txtData.Text), lTipoCombustivel, lNota, lCodigoFornecedor, lOrdem) Then
                Call SubtraiEstoque(lTipoCombustivelAnt, lQuantidadeAnt)
                Call IncluiIntegracaoNuvem(True, True)
                If g_nome_usuario = "L.M.C." Then
                    If (MsgBox("Deseja Excluir TODAS as NF de: " & Mid(txtData.Text, 4, 7) & " ?", vbYesNo + vbQuestion + vbDefaultButton2, "Exclusão MENSAL de Registros!")) = vbYes Then
                        Call GravaAuditoria(1, Me.name, 10, "Exclusão de todas NF de Mes/Ano:" & Mid(txtData.Text, 4, 7))
                        If EntradaCombustivel.ExcluirMesAno(g_empresa, CDate(txtData.Text)) Then
                            MsgBox "As NF de: " & Mid(txtData.Text, 4, 7) & " foram excluídas com sucesso.", vbInformation, "Exclusão MENSAL Concluída!"
                        Else
                            MsgBox "Não foi possível excluir as NF de: " & Mid(txtData.Text, 4, 7) & ".", vbInformation, "Erro de Integridade!"
                        End If
                    End If
                End If
            Else
                MsgBox "Não foi possível excluir este registro!", vbInformation, "Erro de Integridade!"
            End If
            LimpaTela
            If EntradaCombustivel.LocalizarUltimo(g_empresa) Then
                AtualizaTela
            Else
                DesativaBotoes
                cmd_novo.Enabled = True
                cmd_sair.Enabled = True
                cmd_novo.SetFocus
            End If
        End If
    End If
End Sub
Private Sub cmd_novo_Click()
    Call GravaAuditoria(1, Me.name, 2, "")
    LimpaTela
    Inclui
    frm_dados.Enabled = True
    DesabilitaCamposTributacao
    txtDataEmissao.SetFocus

'    If Mid(cboCst.Text, 1, 3) = "060" Then
'        txtValorBcIcmsSt.Enabled = False
'        txtValorIcmsSt.Enabled = False
'        txtValorNaoTributadoRedBcIcms.Enabled = False
'    End If
End Sub
Private Sub cmd_ok_Click()
    On Error GoTo FileError
    If ValidaCampos Then
        AtivaBotoes
        If lOpcao = 1 Then
            Call GravaAuditoria(1, Me.name, 10, "Data:" & txtData.Text & " Forn:" & Val(dtcbo_fornecedor.BoundText) & " NF:" & txt_nota.Text & " Comb:" & Mid(cbo_combustivel.Text, 1, 2) & " Qtd:" & txtQuantidade.Text & " Total:" & txtValorTotal.Text)
            AtualizaTabela
            If EntradaCombustivel.Incluir(False) Then
                lData = CDate(txtData.Text)
                lNota = txt_nota.Text
                lTipoCombustivel = Mid(cbo_combustivel.Text, 1, 2)
                lCodigoFornecedor = Val(dtcbo_fornecedor.BoundText)
                lItem = Val(txtItem.Text)
                lOrdem = EntradaCombustivel.Ordem
                lTipoCombustivelAnt = Mid(cbo_combustivel.Text, 1, 2)
                lQuantidadeAnt = fValidaValor(txtQuantidade.Text)
                Call AdicionaEstoque(lTipoCombustivelAnt, lQuantidadeAnt)
                Call IncluiIntegracaoNuvem(False, False)
                AtualizaPrecoBicos
            Else
                MsgBox "Não foi possível incluir este registro!", vbInformation, "Erro de Integridade!"
            End If
            cmd_novo.SetFocus
        ElseIf lOpcao = 2 Then
            Call GravaAuditoria(1, Me.name, 10, "De Data:" & lData & " Forn:" & lCodigoFornecedorAnt & " NF:" & lNota & " Comb:" & lTipoCombustivelAnt & " Qtd:" & lQuantidadeAnt & " Total:" & lTotalAnt)
            Call GravaAuditoria(1, Me.name, 10, "Para Data:" & txtData.Text & " Forn:" & Val(dtcbo_fornecedor.BoundText) & " NF:" & txt_nota.Text & " Comb:" & Mid(cbo_combustivel.Text, 1, 2) & " Qtd:" & txtQuantidade.Text & " Total:" & txtValorTotal.Text)
            Call SubtraiEstoque(lTipoCombustivelAnt, lQuantidadeAnt)
            AtualizaTabela
            If EntradaCombustivel.Alterar(g_empresa, lData, lTipoCombustivel, lNota, lCodigoFornecedor, lOrdem) Then
                Call AdicionaEstoque(lTipoCombustivelAnt, lQuantidadeAnt)
                Call IncluiIntegracaoNuvem(True, False)
                lData = CDate(txtData.Text)
                lNota = txt_nota.Text
                lTipoCombustivel = Mid(cbo_combustivel.Text, 1, 2)
                lCodigoFornecedor = Val(dtcbo_fornecedor.BoundText)
                lItem = Val(txtItem.Text)
                lOrdem = EntradaCombustivel.Ordem
                lTipoCombustivelAnt = Mid(cbo_combustivel.Text, 1, 2)
                lQuantidadeAnt = fValidaValor(txtQuantidade.Text)
                AtualizaPrecoBicos
            Else
                MsgBox "Não foi possível alterar este registro!", vbInformation, "Erro de Integridade!"
            End If
            
            cmd_novo.SetFocus
        End If
        If EntradaCombustivel.LocalizarCodigo(g_empresa, lData, lTipoCombustivel, lNota, lOrdem) Then
            AtualizaTela
        End If
    End If
    Exit Sub
FileError:
    Exit Sub
End Sub
Function ValidaCampos() As Boolean
    ValidaCampos = False
    If Not IsDate(txtData.Text) Then
        MsgBox "Informe a Data da Entrada.", vbInformation, "Atenção!"
        txtData.SetFocus
    ElseIf Not Val(dtcbo_fornecedor.BoundText) > 0 Then
        MsgBox "Selecione o fornecedor.", vbInformation, "Atenção!"
        dtcbo_fornecedor.SetFocus
    ElseIf txtModelo.Text = "55" And Len(txtChaveAcesso.Text) < 44 Then
        MsgBox "A chave de acesso deve ter 44 numeros.", vbInformation, "Atenção!"
        txtChaveAcesso.SetFocus
    ElseIf txt_nota.Text = "" Then
        MsgBox "Informe o número da Nota Fiscal.", vbInformation, "Atenção!"
        txt_nota.SetFocus
    ElseIf cboFormaPagamento.ListIndex = -1 Then
        MsgBox "Selecione uma forma de pagamento.", vbInformation, "Atenção!"
        cboFormaPagamento.SetFocus
    ElseIf cbo_transporte.ListIndex = -1 Then
        MsgBox "Selecione o tipo de transporte.", vbInformation, "Atenção!"
        cbo_transporte.SetFocus
    ElseIf cbo_combustivel.ListIndex = -1 Then
        MsgBox "Informe o Combustível.", vbInformation, "Atenção!"
        cbo_combustivel.SetFocus
    ElseIf txtValorLitro.Text = "" Then
        MsgBox "Informe o Valor do Litro.", vbInformation, "Atenção!"
        txtValorLitro.SetFocus
    ElseIf Val(txtQuantidade.Text) = 0 Then
        MsgBox "Informe a Quantidade de Entrada.", vbInformation, "Atenção!"
        txtQuantidade.SetFocus
    ElseIf Val(txtItem.Text) = 0 Then
        MsgBox "Informe o ítem da Entrada.", vbInformation, "Atenção!"
        txtItem.SetFocus
    ElseIf txtModelo.Text = "" Then
        MsgBox "Informe o Modelo da Nota Fiscal.", vbInformation, "Atenção!"
        txtModelo.SetFocus
    ElseIf txtSerie.Text = "" Then
        MsgBox "Informe a série da Nota Fiscal.", vbInformation, "Atenção!"
        txtSerie.SetFocus
    ElseIf Not ChamaValidaChaveAcessoNFe Then
        txtChaveAcesso.SetFocus
    ElseIf Not ValidaCST Then
        cboCst.SetFocus
    ElseIf Not ValidaDataEntrada Then
        txtData.SetFocus
    ElseIf Not IsDate(txtData.Text) Then
        MsgBox "Informe a Data da Emissão.", vbInformation, "Atenção!"
        txtDataEmissao.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Function ValidaDataEntrada() As Boolean
    ValidaDataEntrada = False
    
    If LivroLMC.LocalizarCombustivelConcluido(g_empresa, Mid(cbo_combustivel.Text, 1, 2), CDate(txtData.Text)) = "SIM" Then
        MsgBox "O LMC referente a esta data já foi fechado. Por esse motivo não será permitido incluir esta nota", vbInformation, "Atenção!"
        txtData.SetFocus
    Else
        ValidaDataEntrada = True
    End If
End Function
Private Function ValidaChaveAcessoNFe(ByVal pChaveAcesso As String, ByVal pUfEmitente As String, ByVal pDataEntrada As Date, ByVal pCnpjEmitente As String, ByVal pSerie As String, ByVal pNumeroNF As Long) As String
    Dim xDigito As Integer
    Dim i As Integer
    Dim c As Byte
    Dim Key As Integer
    
    ValidaChaveAcessoNFe = "Erro Nao Identificado"
    
    'Verifica Tamanho
    If Len(pChaveAcesso) <> 44 Then
        ValidaChaveAcessoNFe = "Tamanho inferior a 44 digitos"
        Exit Function
    End If
    
    'Verifica código da UF
    If UnidadeFederacao.LocalizarCodigoIBGE(Val(Mid(pChaveAcesso, 1, 2))) Then
        If UnidadeFederacao.Codigo <> pUfEmitente Then
            ValidaChaveAcessoNFe = "UF divergente com o emitente"
            Exit Function
        End If
    End If
    
    'Verifica Ano
    If Mid(pChaveAcesso, 3, 2) <> Mid(Format(pDataEntrada, "dd/MM/yyyy"), 9, 2) Then
        ValidaChaveAcessoNFe = "Ano divergente"
        Exit Function
    End If
    
    'Verifica Mes
    If Mid(pChaveAcesso, 5, 2) <> Mid(Format(pDataEntrada, "dd/MM/yyyy"), 4, 2) Then
        ValidaChaveAcessoNFe = "Mês divergente"
        Exit Function
    End If
    
    'Verifica CNPJ
    If Mid(pChaveAcesso, 7, 14) <> pCnpjEmitente Then
        ValidaChaveAcessoNFe = "CNPJ divergente"
        Exit Function
    End If
    
    'Verifica Modelo do Documento Fiscal
    If Mid(pChaveAcesso, 21, 2) <> "55" Then
        ValidaChaveAcessoNFe = "Modelo do Documento Fiscal divergente"
        Exit Function
    End If
    
    'Verifica Serie do Documento Fiscal
    If Val(Mid(pChaveAcesso, 23, 3)) <> Val(pSerie) Then
        ValidaChaveAcessoNFe = "Série do Documento Fiscal divergente"
        Exit Function
    End If
    
    'Verifica Número do Documento Fiscal
    If CLng(Mid(pChaveAcesso, 26, 9)) <> pNumeroNF Then
        ValidaChaveAcessoNFe = "Número da Nota Fiscal divergente"
        Exit Function
    End If
    
    'Verifica Forma de Emissão do Documento Fiscal
    If Val(Mid(pChaveAcesso, 35, 1)) < 1 Or Val(Mid(pChaveAcesso, 26, 1)) > 5 Then
        ValidaChaveAcessoNFe = "Forma de emissão divergente"
        Exit Function
    End If
    
    'Verifica Dígito Verificador
    c = 2
    'faz um loop por cada numero o mutiplicando-o pelos valores de C
    For i = (Len(pChaveAcesso) - 1) To 1 Step -1
        'vericica se o valor de c for maior que nove,
        'passa o valor para 2
        If c > 9 Then
            c = 2
        End If
        'soma os valores mutiplicados
        Key = Key + (CInt(Mid(pChaveAcesso, i, 1)) * c)
        c = c + 1
    Next
    'obtem o Digito Verificador
    If (Key Mod 11) = 0 Or (Key Mod 11) = 1 Then
        xDigito = 0
    Else
        xDigito = 11 - (Key Mod 11)
    End If
    If Val(Mid(pChaveAcesso, 44, 1)) <> xDigito Then
        ValidaChaveAcessoNFe = "Dígito divergente"
        Exit Function
    End If
    
    'Verifica Modelo do Documento
    If txtModelo.Text <> "55" Then
        ValidaChaveAcessoNFe = "Modelo do documento deve ser 55"
        Exit Function
    End If
   
    ValidaChaveAcessoNFe = "OK"
End Function
Function ValidaCST() As Boolean
    Dim xCST As String
    Dim i As Integer
    Dim xCstErrada As Boolean
    
    ValidaCST = False
    xCST = "060"
    xCstErrada = False
    If Mid(cbo_combustivel.Text, 1, 1) = "A" And lCstAlcool060 = False Then
        If UCase(dtcbo_fornecedor.Text) Like "*ELLO*" Or VerificaFornecedorCst010 Then
            xCST = "010"
            If Mid(cboCst.Text, 1, 3) <> "010" Then
                xCstErrada = True
                MsgBox "A CST para o combustível selecionado deveria ser 010.", vbInformation, "CST selecionada está incorreta!"
                cboCst.SetFocus
            Else
                ValidaCST = True
            End If
        Else
            xCST = "070"
            If Mid(cboCst.Text, 1, 3) <> "070" Then
                xCstErrada = True
                MsgBox "A CST para o combustível selecionado deveria ser 070.", vbInformation, "CST selecionada está incorreta!"
                cboCst.SetFocus
            Else
                If fValidaValor(txtValorNaoTributadoRedBcIcms.Text) = 0 Then
                    MsgBox "Favor informar o Valor Não Tributado da Redução da Base de Cálculo do ICMS." & vbCrLf & "Para calcular o valor, siga a fórmula abaixo:" & vbCrLf & vbCrLf & "(Total de Produtos) - (SOMA do Total de Outros Combustíveis) - (Base de Cálculo do ICMS)", vbInformation, "Campo Obrigatório para Este Combustível!"
                    txtValorNaoTributadoRedBcIcms.SetFocus
                Else
                    ValidaCST = True
                End If
            End If
        End If
    Else
        If Mid(cboCst.Text, 1, 3) <> "060" Then
            xCstErrada = True
            MsgBox "A CST para o combustível selecionado deveria ser 060.", vbInformation, "CST selecionada está incorreta!"
            cboCst.SetFocus
        Else
            ValidaCST = True
        End If
    End If
            
    If xCstErrada Then
        If MsgBox("Deseja que o sistema corrija a CST para correta?", vbYesNo + vbQuestion + vbDefaultButton1, "Correção da CST!") = vbYes Then
            For i = 0 To cboCst.ListCount - 1
                cboCst.ListIndex = i
                If Mid(cboCst.Text, 1, 3) = xCST Then
                    Exit For
                End If
            Next
        End If
    End If
End Function
Private Function VerificaLiberacaoLMC(ByVal pTipoCombustivel As String, ByVal pData As Date) As Boolean
    'If g_nome_usuario = "L.M.C." Then
        VerificaLiberacaoLMC = False
    
        If LivroLMC.LocalizarCombustivelConcluido(g_empresa, pTipoCombustivel, pData) = "NAO" Then
            VerificaLiberacaoLMC = True
            cmd_alterar.Enabled = True
            cmd_excluir.Enabled = True
            
'        Else  'old 01/09 daqui...
'            cmd_alterar.Enabled = False
'            cmd_excluir.Enabled = False
'        End If '... ate aqui old 01/09

        Else 'new 01/09 daqui...
            If LivroLMC.LocalizarCombustivelConcluido(g_empresa, pTipoCombustivel, pData) = "SIM" Then
                cmd_alterar.Enabled = False
                cmd_excluir.Enabled = False
            Else
                cmd_alterar.Enabled = True
                cmd_excluir.Enabled = True
            End If
        End If '...ate aqui new 01/09
    'Else
    '    VerificaLiberacaoLMC = True
    'End If
End Function
Private Function VerificaFornecedorCst010() As Boolean
    Dim i As Integer
    
    VerificaFornecedorCst010 = False
    For i = 1 To 10
        If ConfiguracaoDiversa.LocalizarCodigo(1, "FORNECEDOR ALCOOL CST 010 SEQ:" & Format(i, "00")) Then
            If dtcbo_fornecedor.Text Like "*" & ConfiguracaoDiversa.Texto & "*" Then
                VerificaFornecedorCst010 = True
                Exit For
            End If
        End If
    Next
End Function
Private Sub PreparaCstAlcool060()
    lCstAlcool060 = False
    If ConfiguracaoDiversa.LocalizarCodigo(1, "ALCOOL CST 060") Then
        If ConfiguracaoDiversa.Verdadeiro Then
            lCstAlcool060 = True
        End If
    End If
End Sub
Private Sub SubtraiEstoque(x_tipo_combustivel As String, x_quantidade As Currency)
    If Combustivel.LocalizarCodigo(g_empresa, x_tipo_combustivel) Then
        Combustivel.QuantidadeEmEstoque = Combustivel.QuantidadeEmEstoque - x_quantidade
        If Not Combustivel.Alterar(g_empresa, x_tipo_combustivel) Then
            MsgBox "Não foi possível subtrair no estoque", vbInformation, "Registro Não Encontrado"
        End If
    End If
End Sub
Private Sub TransfereDadosLMC()
'    Dim x_data As Date
'
'    'Busca ultima data com movimento
'    x_data = CDate("01/01/1900")
'    If EntradaCombustivel.LocalizarUltimo(g_empresa) Then
'        x_data = EntradaCombustivel.Data
'    End If
'    x_data = x_data + 1
'    If (MsgBox("Na empresa " & g_nome_empresa & Chr(10) & "Será transferido a entrada de combustível apartir da data " & x_data & "." & Chr(10) & Chr(10) & "Deseja realmente fazer esta transferência?", vbYesNo + 256, "Transfere a Entrada de Combustível Para o L.M.C.!")) = 7 Then
'        Exit Sub
'    End If
'    Call GravaAuditoria(1, Me.name, 10, "Empresa:" & g_empresa & " A Partir de:" & x_data)
'
'    'Transfere Dados para o LMC
'    If EntradaCombustivel.TransfereDados(g_empresa, x_data, "Entrada_Combustivel") Then
'        MsgBox "A empresa " & g_nome_empresa & Chr(10) & "Está com a entrada de combustível transferida para o L.M.C.", vbInformation, "Transferência Concluida!"
'    Else
'        MsgBox "A empresa " & g_nome_empresa & Chr(10) & "Não tem entrada de combustível à ser transferida para o L.M.C.", vbInformation, "Transferência Não Concluida!"
'    End If
End Sub
Private Sub cmd_pesquisa_Click()
    Call GravaAuditoria(1, Me.name, 5, "")
    consulta_movimento_entrada_combustivel.Show 1
    If Len(g_string) > 0 Then
        lData = RetiraGString(1)
        lTipoCombustivel = RetiraGString(2)
        lNota = RetiraGString(3)
        lOrdem = RetiraGString(4)
        If EntradaCombustivel.LocalizarCodigo(g_empresa, lData, lTipoCombustivel, lNota, lOrdem) Then
            AtualizaTela
        End If
    End If
End Sub
Private Sub cmd_primeiro_Click()
    Call GravaAuditoria(1, Me.name, 12, "")
    If EntradaCombustivel.LocalizarPrimeiro Then
        AtualizaTela
        cmd_proximo.SetFocus
    Else
        MsgBox "Não há registros nesta empresa.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    Call GravaAuditoria(1, Me.name, 14, "")
    If EntradaCombustivel.LocalizarProximo Then
        AtualizaTela
    Else
        MsgBox "Fim de Arquivo.", vbInformation, "Atenção!"
        cmd_anterior.SetFocus
    End If
End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub cmd_transfere_dados_lmc_Click()
    Call GravaAuditoria(1, Me.name, 23, "Transferencia Para LMC")
    If EntradaCombustivel.TransfereDadosLMC(g_empresa, True) Then
        Call GravaAuditoria(1, Me.name, 10, "Empresa:" & g_empresa & " A Partir de:" & EntradaCombustivel.UltimaData(g_empresa))
        If EntradaCombustivel.TransfereDadosLMC(g_empresa, False) Then
            MsgBox "A empresa " & g_nome_empresa & Chr(10) & "Está com a entrada de combustível transferida para o L.M.C.", vbInformation, "Transferência Concluida!"
        Else
            MsgBox "A empresa " & g_nome_empresa & Chr(10) & "Não tem entrada de combustível à ser transferida para o L.M.C.", vbInformation, "Transferência Não Concluida!"
        End If
    Else
        MsgBox "A empresa " & g_nome_empresa & Chr(10) & "Não tem entrada de combustível à ser transferida para o L.M.C.", vbInformation, "Transferência Não Concluida!"
    End If
    'TransfereDadosLMC
    cmd_cancelar_Click
    cmd_ultimo_Click
End Sub
Private Sub cmd_ultimo_Click()
    Call GravaAuditoria(1, Me.name, 15, "")
    If EntradaCombustivel.LocalizarUltimo(g_empresa) Then
        AtualizaTela
        cmd_anterior.SetFocus
    Else
        MsgBox "Não há registros nesta empresa.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub dtcbo_fornecedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txt_nota.SetFocus
    End If
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    If flag_movimento_Entrada_combustivel = 0 Then
        Set adodc_fornecedor.Recordset = Conectar.RsConexao("SELECT Codigo, Nome FROM Fornecedor WHERE Empresa = " & g_empresa & " ORDER BY Nome")
        DesativaBotoes
        If EntradaCombustivel.LocalizarUltimo(g_empresa) Then
            AtivaBotoes
            AtualizaTela
            Call VerificaLiberacaoLMC(lTipoCombustivelAnt, lData)
        Else
            cmd_novo.Enabled = True
            cmd_sair.Enabled = True
        End If
        If cmd_novo.Enabled Then
            cmd_novo.SetFocus
        End If
    Else
        flag_movimento_Entrada_combustivel = 0
    End If
End Sub
Private Sub Form_Deactivate()
    flag_movimento_Entrada_combustivel = 1
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
    Screen.MousePointer = 1
    CentraForm Me
    
    If g_nome_usuario = "L.M.C." Then
        EntradaCombustivel.NomeTabela = "Entrada_Combustivel_LMC"
        Me.Caption = Me.Caption & " - LMC"
    Else
        EntradaCombustivel.NomeTabela = "Entrada_Combustivel"
    End If
    PreencheCboCst
    PreencheCboCombustivel
    PreencheCboTransporte
    PreencheCboFormaPagamento
    PreparaCstAlcool060
End Sub
Private Sub PreencheCboCombustivel()
    Dim xSQL As String
    
    cbo_combustivel.Clear
    'Prepara SQL
    xSQL = "SELECT Nome, Codigo"
    xSQL = xSQL & "  FROM Combustivel"
    xSQL = xSQL & " WHERE Empresa = " & g_empresa
    xSQL = xSQL & " ORDER BY Nome"
    'Abre RecordSet
    Set rsTabela = New adodb.Recordset
    Set rsTabela = Conectar.RsConexao(xSQL)
    'Verifica tabela
    If rsTabela.RecordCount > 0 Then
        Do Until rsTabela.EOF
            cbo_combustivel.AddItem rsTabela("Codigo").Value & " - " & rsTabela("Nome").Value
            rsTabela.MoveNext
        Loop
    End If
    rsTabela.Close
    Set rsTabela = Nothing
End Sub
Private Sub PreencheCboCst()
    cboCst.Clear
    cboCst.AddItem "010-Tributada c/ICMS p/ST"
    cboCst.ItemData(cboCst.NewIndex) = 0
    cboCst.AddItem "060-ICMS Cobrado Int. p/ST"
    cboCst.ItemData(cboCst.NewIndex) = 1
    cboCst.AddItem "070-ICMS Cobrado c/Red.BC p/ST"
    cboCst.ItemData(cboCst.NewIndex) = 2
End Sub
Private Sub PreencheCboFormaPagamento()
    cboFormaPagamento.Clear
    cboFormaPagamento.AddItem "0-À Vista"
    cboFormaPagamento.ItemData(cboFormaPagamento.NewIndex) = 0
    cboFormaPagamento.AddItem "1-À Prazo"
    cboFormaPagamento.ItemData(cboFormaPagamento.NewIndex) = 1
    cboFormaPagamento.AddItem "2-Sem Pagamento"
    cboFormaPagamento.ItemData(cboFormaPagamento.NewIndex) = 2
End Sub
Private Sub PreencheCboTanque()
    Dim xSQL As String
    
    cboTanque.Clear
    'Prepara SQL
    xSQL = "SELECT [Numero do Tanque] AS Codigo"
    xSQL = xSQL & "  FROM Tanque_Combustivel"
    xSQL = xSQL & " WHERE Empresa = " & g_empresa
    xSQL = xSQL & "   AND [Tipo de Combustivel] = " & preparaTexto(Mid(cbo_combustivel.Text, 1, 2))
    xSQL = xSQL & " ORDER BY [Numero do Tanque]"
    'Abre RecordSet
    Set rsTabela = New adodb.Recordset
    Set rsTabela = Conectar.RsConexao(xSQL)
    'Verifica tabela
    If rsTabela.RecordCount > 0 Then
        Do Until rsTabela.EOF
            cboTanque.AddItem rsTabela("Codigo").Value
            rsTabela.MoveNext
        Loop
    End If
    rsTabela.Close
    Set rsTabela = Nothing
    
    If cboTanque.ListCount > 0 Then
        cboTanque.ListIndex = 0
    End If
End Sub
Private Sub PreencheCboTransporte()
    cbo_transporte.Clear
    cbo_transporte.AddItem "0-Terceiros"
    cbo_transporte.ItemData(cbo_transporte.NewIndex) = 0
    cbo_transporte.AddItem "1-Emitente"
    cbo_transporte.ItemData(cbo_transporte.NewIndex) = 1
    cbo_transporte.AddItem "2-Destinatário"
    cbo_transporte.ItemData(cbo_transporte.NewIndex) = 2
    cbo_transporte.AddItem "9-Sem Cobrança"
    cbo_transporte.ItemData(cbo_transporte.NewIndex) = 9
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub txt_nota_LostFocus()
    If Len(txt_nota.Text) > 0 Then
        txt_nota.Text = Val(txt_nota.Text)
    End If
End Sub

Private Sub txtAliquotaIcms_GotFocus()
    txtAliquotaIcms.SelStart = 0
    txtAliquotaIcms.SelLength = Len(txtAliquotaIcms.Text)
End Sub
Private Sub txtAliquotaIcms_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txtValorIcms.SetFocus
    End If
End Sub
Private Sub txtAliquotaIcms_LostFocus()
    If txtAliquotaIcms.Text = "" Then
        txtAliquotaIcms.Text = "0"
    End If
    txtAliquotaIcms.Text = Format(txtAliquotaIcms.Text, "###,###,##0.00")
    If fValidaValor(txtAliquotaIcms.Text) > 0 And fValidaValor(txtValorBcIcms.Text) > 0 Then
        txtValorIcms.Text = Format(fValidaValor(txtValorBcIcms.Text) * fValidaValor(txtAliquotaIcms.Text) / 100, "##,###,##0.00")
    End If
End Sub
Private Sub txtChaveAcesso_GotFocus()
    txtChaveAcesso.SelStart = 0
    txtChaveAcesso.SelLength = Len(txtChaveAcesso.Text)
End Sub
Private Sub txtChaveAcesso_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cboFormaPagamento.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txtItem_GotFocus()
    txtItem.SelStart = 0
    txtItem.SelLength = Len(txtItem.Text)
End Sub
Private Sub txtItem_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cboTanque.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txtItem_LostFocus()
    txtItem.Text = Format(Val(txtItem.Text), "00")
End Sub
Private Sub txtModelo_GotFocus()
    txtModelo.SelStart = 0
    txtModelo.SelLength = Len(txtModelo.Text)
End Sub
Private Sub txtModelo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtSerie.SetFocus
    End If
End Sub
Private Sub txtQuantidade_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txtValorTotal.SetFocus
    End If
    Call ValidaValorSinal(KeyAscii)
End Sub
Private Sub txtQuantidade_LostFocus()
    If txtQuantidade.Text = "" Then
        txtQuantidade.Text = "0"
    End If
    txtQuantidade.Text = Format(txtQuantidade.Text, "###,##0.0")
    If fValidaValor(txtQuantidade.Text) > 0 And fValidaValor(txtValorLitro.Text) > 0 Then
        txtValorTotal.Text = Format((fValidaValor(txtQuantidade.Text) * fValidaValor(txtValorLitro.Text)), "###,###,##0.00")
    End If
End Sub
Private Sub txtSerie_GotFocus()
    txtSerie.SelStart = 0
    txtSerie.SelLength = Len(txtSerie.Text)
End Sub
Private Sub txtSerie_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtChaveAcesso.SetFocus
    End If
End Sub
Private Sub txtValorBcIcms_GotFocus()
    txtValorBcIcms.SelStart = 0
    txtValorBcIcms.SelLength = Len(txtValorBcIcms.Text)
End Sub
Private Sub txtValorBcIcms_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txtAliquotaIcms.SetFocus
    End If
End Sub
Private Sub txtValorBcIcms_LostFocus()
    If txtValorBcIcms.Text = "" Then
        txtValorBcIcms.Text = "0"
    End If
    txtValorBcIcms.Text = Format(txtValorBcIcms.Text, "###,###,##0.00")
End Sub
Private Sub txtValorBcIcmsSt_GotFocus()
    txtValorBcIcmsSt.SelStart = 0
    txtValorBcIcmsSt.SelLength = Len(txtValorBcIcmsSt.Text)
End Sub
Private Sub txtValorBcIcmsSt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txtValorIcmsSt.SetFocus
    End If
End Sub
Private Sub txtValorBcIcmsSt_LostFocus()
    If txtValorBcIcmsSt.Text = "" Then
        txtValorBcIcmsSt.Text = "0"
    End If
    txtValorBcIcmsSt.Text = Format(txtValorBcIcmsSt.Text, "###,###,##0.00")
End Sub
Private Sub txtValorFrete_GotFocus()
    txtValorFrete.SelStart = 0
    txtValorFrete.SelLength = Len(txtValorFrete.Text)
End Sub
Private Sub txtValorFrete_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub txtValorFrete_LostFocus()
    txtValorFrete.Text = Format(txtValorFrete.Text, "###,###,##0.00")
    If txtValorFrete.Text <> "" Then
        txtValorFrete.Text = Format(txtValorFrete, "###,###,##0.00")
    End If
End Sub

Private Sub txtValorIcms_GotFocus()
    txtValorIcms.SelStart = 0
    txtValorIcms.SelLength = Len(txtValorIcms.Text)
End Sub
Private Sub txtValorIcms_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        cboCst.SetFocus
    End If
End Sub
Private Sub txtValorIcms_LostFocus()
    If txtValorIcms.Text = "" Then
        txtValorIcms.Text = "0"
    End If
    txtValorIcms.Text = Format(txtValorIcms.Text, "###,###,##0.00")
End Sub
Private Sub txtValorIcmsSt_GotFocus()
    txtValorIcmsSt.SelStart = 0
    txtValorIcmsSt.SelLength = Len(txtValorIcmsSt.Text)
End Sub
Private Sub txtValorIcmsSt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txtValorNaoTributadoRedBcIcms.SetFocus
    End If
End Sub
Private Sub txtValorIcmsSt_LostFocus()
    If txtValorIcmsSt.Text = "" Then
        txtValorIcmsSt.Text = "0"
    End If
    txtValorIcmsSt.Text = Format(txtValorIcmsSt.Text, "###,###,##0.00")
End Sub
Private Sub txtValorLitro_GotFocus()
    txtValorLitro.SelStart = 0
    txtValorLitro.SelLength = Len(txtValorLitro.Text)
End Sub
Private Sub txtValorLitro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txtQuantidade.SetFocus
    End If
End Sub
Private Sub txtValorLitro_LostFocus()
    If txtValorLitro.Text = "" Then
        txtValorLitro.Text = "0"
    End If
    txtValorLitro.Text = Format(txtValorLitro.Text, "###,##0.0000")
End Sub
Private Sub txt_nota_GotFocus()
    txt_nota.SelStart = 0
    txt_nota.SelLength = Len(txt_nota.Text)
End Sub
Private Sub txt_nota_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Then
        KeyAscii = 46
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        cbo_transporte.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txtValorNaoTributadoRedBcIcms_GotFocus()
    txtValorNaoTributadoRedBcIcms.SelStart = 0
    txtValorNaoTributadoRedBcIcms.SelLength = Len(txtValorNaoTributadoRedBcIcms.Text)
End Sub
Private Sub txtValorNaoTributadoRedBcIcms_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txtValorFrete.SetFocus
    End If
End Sub
Private Sub txtValorNaoTributadoRedBcIcms_LostFocus()
    If txtValorNaoTributadoRedBcIcms.Text = "" Then
        txtValorNaoTributadoRedBcIcms.Text = "0"
    End If
    txtValorNaoTributadoRedBcIcms.Text = Format(txtValorNaoTributadoRedBcIcms.Text, "###,###,##0.00")
End Sub
Private Sub txtValorTotal_GotFocus()
    txtValorTotal.SelStart = 0
    txtValorTotal.SelLength = Len(txtValorTotal.Text)
End Sub
Private Sub txtValorTotal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        If Mid(cboCst.Text, 1, 3) = "060" Then
            KeyAscii = 0
            cboCst.SetFocus
        Else
            DesabilitaCamposTributacao
            KeyAscii = 0
            txtValorBcIcms.SetFocus
        End If
        
    End If
End Sub
Private Sub txtValorTotal_LostFocus()
    If txtValorTotal.Text = "" Then
        txtValorTotal.Text = "0"
    End If
    txtValorTotal.Text = Format(txtValorTotal.Text, "###,###,##0.00")
    If fValidaValor(txtQuantidade.Text) > 0 Then
        txtValorLitro.Text = Format(fValidaValor(txtValorTotal.Text) / fValidaValor(txtQuantidade.Text), "###,##0.0000")
    End If
End Sub
Private Sub txtData_GotFocus()
    If Not IsDate(txtData.Text) Then
        txtData.Text = Format(txtDataEmissao, "dd/mm/yyyy")
    End If
    txtData.Text = fDesmascaraData(txtData.Text)
    txtData.SelStart = 0
    txtData.SelLength = 4
    txtData.MaxLength = 8
End Sub
Private Sub txtData_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtModelo.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txtData_LostFocus()
    txtData.MaxLength = 10
    txtData.Text = fMascaraData(txtData.Text)
End Sub
Private Sub txtDataEmissao_GotFocus()
    If Not IsDate(txtDataEmissao.Text) Then
        If IsDate(lData) And lData <> "00:00:00" Then
            txtDataEmissao.Text = Format(lData, "dd/mm/yyyy")
        Else
            txtDataEmissao.Text = Format(g_data_def, "dd/mm/yyyy")
        End If
        txt_nota.Text = lNota
        txtChaveAcesso.Text = lChaveAcesso
        dtcbo_fornecedor.BoundText = EntradaCombustivel.RetornaUltimoFornecedor(g_empresa)
        cbo_transporte.ListIndex = 0
    End If
    txtDataEmissao.Text = fDesmascaraData(txtDataEmissao.Text)
    txtDataEmissao.SelStart = 0
    txtDataEmissao.SelLength = 4
    txtDataEmissao.MaxLength = 8
End Sub
Private Sub txtDataEmissao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtData.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txtDataEmissao_LostFocus()
    txtDataEmissao.MaxLength = 10
    txtDataEmissao.Text = fMascaraData(txtDataEmissao.Text)
End Sub



