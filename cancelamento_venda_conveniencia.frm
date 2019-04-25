VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form cancelamento_venda_conveniencia 
   Caption         =   "Cancelamento de Pedido de Compra"
   ClientHeight    =   7215
   ClientLeft      =   1410
   ClientTop       =   1545
   ClientWidth     =   10950
   Icon            =   "cancelamento_venda_conveniencia.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   Picture         =   "cancelamento_venda_conveniencia.frx":030A
   ScaleHeight     =   7215
   ScaleWidth      =   10950
   Begin VB.Frame frmDados 
      Height          =   7035
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   10815
      Begin VB.CommandButton cmd_sair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   855
         Left            =   9960
         Picture         =   "cancelamento_venda_conveniencia.frx":0750
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Sai e fecha esta janela."
         Top             =   180
         Width           =   795
      End
      Begin VB.TextBox txt_celula 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   540
         TabIndex        =   8
         Top             =   2700
         Visible         =   0   'False
         Width           =   1575
      End
      Begin MSAdodcLib.Adodc adodc_venda_conveniencia 
         Height          =   330
         Left            =   1260
         Top             =   3660
         Visible         =   0   'False
         Width           =   3195
         _ExtentX        =   5636
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
         Caption         =   "adodc_venda_conveniencia"
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
      Begin VB.ComboBox cbo_periodo 
         Height          =   315
         Left            =   5520
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   420
         Width           =   495
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2640
         Picture         =   "cancelamento_venda_conveniencia.frx":1DE2
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   420
         Width           =   495
      End
      Begin MSMask.MaskEdBox msk_data 
         Height          =   300
         Left            =   1500
         TabIndex        =   2
         Top             =   420
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSFlexGridLib.MSFlexGrid fgd_composicao_caixa 
         Height          =   5295
         Left            =   20
         TabIndex        =   7
         Top             =   1080
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   9340
         _Version        =   393216
         Cols            =   10
         FixedCols       =   0
         BackColorBkg    =   12632256
         AllowUserResizing=   1
      End
      Begin VB.Label Label6 
         Caption         =   "Total"
         Height          =   315
         Left            =   7320
         TabIndex        =   9
         Top             =   6540
         Width           =   735
      End
      Begin VB.Label lbl_total 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   8220
         TabIndex        =   10
         Top             =   6480
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "&Data do pedido"
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   1
         Top             =   420
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "&Período"
         Height          =   315
         Index           =   6
         Left            =   4500
         TabIndex        =   4
         Top             =   420
         Width           =   975
      End
   End
End
Attribute VB_Name = "cancelamento_venda_conveniencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lFlagMovimento As Integer
Dim lOpcao As String
Dim lQtdPeriodo As Integer

Dim lEmpresa As Integer
Dim lData As Date
Dim lIlha As Integer
Dim lPeriodo As Integer
Dim lTipoMovimento As Integer
Dim lCodigoComposicao As Integer
Dim lNumeroJustificativa As Long
Dim lCodigoFuncionario As Integer
Dim lNumeroMovimentoCaixa As Long
Dim lOrigemVenda As String
Dim lCxDataDigitacao As Date
Dim lCxHoraDigitacao As Date

Const NovaLinha As String = ">*"      ' Indica uma nova linha
Private ControlVisible As Boolean     ' Se o controle esta visivel ou nao
Private LastRow As Long               ' Ultima linha em que se editou
Private LastCol As Long               ' ultima coluna em que se editou
Dim lMarcaCelula As Boolean

Private rsMovVendaConveniencia As New adodb.Recordset

Private Configuracao As New cConfiguracao
Private Estoque As New cEstoque
Private IntegracaoCaixa As New cIntegracaoCaixa
Private MovimentoCaixaPista As New cMovimentoCaixaPista
Private MovimentoLubrificante As New cMovimentoLubrificante
Private MovimentoVendaConveniencia As New cMovimentoVendaConveniencia
Private MovJustificativa As New cMovimentoJustificativa
Private Produto As New cProduto

Private Sub AtribuiValorCelula()
    Dim Texto As String
    '
    txt_celula.Visible = False
    ControlVisible = False
    '
    ' atribuir o texto anterior a celula
    Select Case LastCol
      Case 4 To 7
        'notas menores que 5 muda cor fonte para vermelho, demais azul
        Texto = txt_celula.Text
        fgd_composicao_caixa.TextMatrix(LastRow, LastCol) = Texto
        'If Val(fgd_composicao_caixa.Text) < 6 Then
        '     fgd_composicao_caixa.CellForeColor = vbRed
        'Else
        '     fgd_composicao_caixa.CellForeColor = vbBlue
        'End If
      Case Else
        'If LastRow = 0 And LastCol = 0 Then
            LastRow = fgd_composicao_caixa.Row
            LastCol = fgd_composicao_caixa.Col
        'End If
      
        Texto = txt_celula.Text
        fgd_composicao_caixa.TextMatrix(LastRow, LastCol) = Texto
    End Select
End Sub
Private Sub AtualizaConstantes()
    If Configuracao.LocalizarCodigo(g_empresa) Then
        lQtdPeriodo = Configuracao.QuantidadePeriodos
    Else
        lQtdPeriodo = 1
    End If
End Sub
Private Sub AtualizaGrid()
    Dim xTotal As Currency
    Dim i As Integer
    Dim xSQL As String
    
    LimpaGrid
    i = 0
    fgd_composicao_caixa.Visible = False
    xTotal = 0
    
    xSQL = ""
    xSQL = xSQL & "   SELECT Data, [Numero do Cupom], Ordem, Periodo, Movimento_Venda_Conveniencia.[Codigo do Produto], Produto.Nome as NomeProduto, Quantidade, [Valor Unitario], ([Valor Total] - [Valor do Desconto]) AS [Valor Total], [Item Cancelado], [Cupom Cancelado]"
    xSQL = xSQL & "     FROM Movimento_Venda_Conveniencia, Produto"
    xSQL = xSQL & "    WHERE Empresa = " & g_empresa
    xSQL = xSQL & "      AND Data = " & preparaData(CDate(msk_data.Text))
    xSQL = xSQL & "      AND Periodo = " & preparaTexto(Val(cbo_periodo.ItemData(cbo_periodo.ListIndex)))
    xSQL = xSQL & "      AND Produto.Codigo = Movimento_Venda_Conveniencia.[Codigo do Produto]"
    xSQL = xSQL & " ORDER BY Data, [Numero do Cupom], Ordem"
    Set rsMovVendaConveniencia = New adodb.Recordset
    Set rsMovVendaConveniencia = Conectar.RsConexao(xSQL)
    If Not rsMovVendaConveniencia.EOF Then
        Do Until rsMovVendaConveniencia.EOF
            i = i + 1
            fgd_composicao_caixa.Rows = fgd_composicao_caixa.Rows + 1
            fgd_composicao_caixa.Row = i
            fgd_composicao_caixa.Col = 0
            fgd_composicao_caixa.Text = Format(rsMovVendaConveniencia("Data").Value, "dd/mm/yyyy")
            fgd_composicao_caixa.Col = 1
            fgd_composicao_caixa.Text = Format(rsMovVendaConveniencia("Numero do Cupom").Value, "###,##0")
            fgd_composicao_caixa.Col = 2
            fgd_composicao_caixa.Text = Format(rsMovVendaConveniencia("Ordem").Value, "###,##0")
            fgd_composicao_caixa.Col = 3
            fgd_composicao_caixa.Text = rsMovVendaConveniencia("Periodo").Value
            fgd_composicao_caixa.Col = 4
            fgd_composicao_caixa.Text = Format(rsMovVendaConveniencia("Codigo do Produto").Value, "###0")
            fgd_composicao_caixa.Col = 5
            fgd_composicao_caixa.Text = rsMovVendaConveniencia("NomeProduto").Value
            fgd_composicao_caixa.Col = 6
            fgd_composicao_caixa.Text = Format(rsMovVendaConveniencia("Valor Unitario").Value, "###,###,##0.00")
            fgd_composicao_caixa.Col = 7
            fgd_composicao_caixa.Text = Format(rsMovVendaConveniencia("Quantidade").Value, "###,###,##0.00")
            fgd_composicao_caixa.Col = 8
            fgd_composicao_caixa.Text = Format(rsMovVendaConveniencia("Valor Total").Value, "###,###,##0.00")
            fgd_composicao_caixa.Col = 9
            If rsMovVendaConveniencia("Item Cancelado").Value Or rsMovVendaConveniencia("Cupom Cancelado").Value Then
                fgd_composicao_caixa.Text = "Cancelado"
            Else
                fgd_composicao_caixa.Text = "Normal"
                xTotal = xTotal + rsMovVendaConveniencia("Valor Total").Value
            End If
            rsMovVendaConveniencia.MoveNext
        Loop
    End If
    
    rsMovVendaConveniencia.Close
    Set rsMovVendaConveniencia = Nothing
    fgd_composicao_caixa.Row = 1
    fgd_composicao_caixa.Col = 9
    fgd_composicao_caixa.Rows = fgd_composicao_caixa.Rows - 1
    fgd_composicao_caixa.Visible = True
    lbl_total.Caption = Format(xTotal, "###,###,##0.00")
    Call GravaAuditoria(1, Me.name, 5, "Total:" & Format(xTotal, "###,###,##0.00") & " Data:" & msk_data.Text & " Per:" & cbo_periodo.ItemData(cbo_periodo.ListIndex))
    fgd_composicao_caixa.SetFocus
End Sub
Private Sub AtualizaTabelaVendaProduto(ByVal pCancelar As Boolean)
    If Not IntegracaoCaixa.LocalizarNome(g_empresa, "VENDA DE LUBRIFICANTES") Then
        MsgBox "Não será possível integrar com o caixa!", vbCritical, "Erro de Integridade!"
    Else
        If pCancelar = False Then
            If IncluiMovimentoCaixa("VENDA DE LUBRIFICANTES") Then
                If MovimentoLubrificante.LocalizarCodigo(g_empresa, MovimentoVendaConveniencia.Data, MovimentoVendaConveniencia.Periodo, MovimentoVendaConveniencia.Ilha, 1, 2, MovimentoVendaConveniencia.CodigoProduto, MovimentoVendaConveniencia.operador) Then
                    MovimentoLubrificante.Quantidade = MovimentoLubrificante.Quantidade + MovimentoVendaConveniencia.Quantidade
                    MovimentoLubrificante.ValorTotal = MovimentoLubrificante.ValorTotal + MovimentoVendaConveniencia.ValorTotal
                    If MovimentoLubrificante.Alterar(g_empresa, MovimentoVendaConveniencia.Data, MovimentoVendaConveniencia.Periodo, MovimentoVendaConveniencia.Ilha, 1, 2, MovimentoVendaConveniencia.CodigoProduto, MovimentoVendaConveniencia.operador) Then
                    Else
                        MsgBox "Não foi possível alterar o registro Venda Produto!", vbCritical, "Erro de Integridade!"
                    End If
                Else
                    MovimentoLubrificante.Empresa = g_empresa
                    MovimentoLubrificante.Data = Format(MovimentoVendaConveniencia.Data, "dd/mm/yyyy")
                    MovimentoLubrificante.Periodo = MovimentoVendaConveniencia.Periodo
                    MovimentoLubrificante.NumeroIlha = MovimentoVendaConveniencia.Ilha
                    MovimentoLubrificante.CodigoTipoSubEstoque = 2
                    MovimentoLubrificante.CodigoFuncionario = MovimentoVendaConveniencia.operador
                    MovimentoLubrificante.CodigoProduto = MovimentoVendaConveniencia.CodigoProduto
                    MovimentoLubrificante.Quantidade = MovimentoVendaConveniencia.Quantidade
                    MovimentoLubrificante.ValorCusto = MovimentoVendaConveniencia.PrecoCusto
                    MovimentoLubrificante.ValorVenda = MovimentoVendaConveniencia.ValorUnitario
                    MovimentoLubrificante.ValorTotal = MovimentoVendaConveniencia.ValorTotal
                    MovimentoLubrificante.OrdemDigitacao = 1
                    MovimentoLubrificante.TipoMovimento = 1
                    If MovimentoLubrificante.Incluir Then
                    Else
                        MsgBox "Não foi possível incluir Venda de Produtos", vbCritical, "Erro de Integridade!"
                    End If
                End If
            Else
                MsgBox "Não foi possível integrar no caixa!", vbCritical, "Erro de Integridade!"
            End If
        ElseIf pCancelar = True Then
            If ExcluiMovimentoCaixa Then
                If MovimentoLubrificante.LocalizarCodigo(g_empresa, MovimentoVendaConveniencia.Data, MovimentoVendaConveniencia.Periodo, MovimentoVendaConveniencia.Ilha, 1, 2, MovimentoVendaConveniencia.CodigoProduto, MovimentoVendaConveniencia.operador) Then
                    If MovimentoLubrificante.Quantidade = MovimentoVendaConveniencia.Quantidade Then
                        If MovimentoLubrificante.Excluir(g_empresa, MovimentoVendaConveniencia.Data, MovimentoVendaConveniencia.Periodo, MovimentoVendaConveniencia.Ilha, 1, 2, MovimentoVendaConveniencia.CodigoProduto, MovimentoVendaConveniencia.operador) Then
                        Else
                            MsgBox "Não foi possível excluir o registro Venda Produto!", vbCritical, "Erro de Integridade!"
                        End If
                    Else
                        MovimentoLubrificante.Quantidade = MovimentoLubrificante.Quantidade - MovimentoVendaConveniencia.Quantidade
                        MovimentoLubrificante.ValorTotal = MovimentoLubrificante.ValorTotal - MovimentoVendaConveniencia.ValorTotal
                        If MovimentoLubrificante.Alterar(g_empresa, MovimentoVendaConveniencia.Data, MovimentoVendaConveniencia.Periodo, MovimentoVendaConveniencia.Ilha, 1, 2, MovimentoVendaConveniencia.CodigoProduto, MovimentoVendaConveniencia.operador) Then
                        Else
                            MsgBox "Não foi possível alterar o registro Venda Produto!", vbCritical, "Erro de Integridade!"
                        End If
                    End If
                Else
                    MsgBox "Não foi possível localizar Venda de Produtos", vbCritical, "Erro de Integridade!"
                End If
            Else
                MsgBox "Não foi possível integrar no caixa!", vbCritical, "Erro de Integridade!"
            End If
        End If
    End If
End Sub
Private Sub AlteraMovimentoVendaConveniencia(ByVal pCancelar As Boolean)
    Dim xOperacao As String
    
    xOperacao = ""
    If MovimentoVendaConveniencia.LocalizarCodigo(g_empresa, fgd_composicao_caixa.TextMatrix(LastRow, 1), fgd_composicao_caixa.TextMatrix(LastRow, 0), lIlha, lOrigemVenda, fgd_composicao_caixa.TextMatrix(LastRow, 2)) Then
        If pCancelar = False Then
            If MovimentoVendaConveniencia.ItemCancelado = True Then
                xOperacao = "-"
            End If
            MovimentoVendaConveniencia.CupomCancelado = False
            MovimentoVendaConveniencia.ItemCancelado = False
            If MovimentoVendaConveniencia.NumeroJustificativa > 0 Then
                If Not MovJustificativa.Excluir(MovimentoVendaConveniencia.NumeroJustificativa) Then
                    MsgBox "Não existe justificativa a ser excluída!", vbInformation, "Erro de Integridade!"
                End If
            End If
            MovimentoVendaConveniencia.NumeroJustificativa = 0
        ElseIf pCancelar = True Then
            If MovimentoVendaConveniencia.ItemCancelado = False Then
                xOperacao = "+"
            End If
            MovimentoVendaConveniencia.ItemCancelado = True
            MovimentoVendaConveniencia.NumeroJustificativa = lNumeroJustificativa
        End If
        If MovimentoVendaConveniencia.Alterar(g_empresa, fgd_composicao_caixa.TextMatrix(LastRow, 1), fgd_composicao_caixa.TextMatrix(LastRow, 0), lIlha, lOrigemVenda, fgd_composicao_caixa.TextMatrix(LastRow, 2)) Then
            If xOperacao = "+" Then
                If Not Estoque.Adicionar(g_empresa, MovimentoVendaConveniencia.CodigoProduto, MovimentoVendaConveniencia.Quantidade) Then
                    MsgBox "Erro ao tentar adicionar no estoque.", vbInformation, "Erro de Integridade!"
                End If
            ElseIf xOperacao = "-" Then
                If Not Estoque.Subtrair(g_empresa, MovimentoVendaConveniencia.CodigoProduto, MovimentoVendaConveniencia.Quantidade) Then
                    MsgBox "Erro ao tentar adicionar no estoque.", vbInformation, "Erro de Integridade!"
                End If
            End If
            Exit Sub
        Else
            MsgBox "Não foi alterado!", vbInformation, "Atenção"
        End If
    Else
        MsgBox "Venda não Encontrada!", vbInformation, "Erro de Integridade"
    End If
    'For i = 1 To (fgd_composicao_caixa.Rows - 2)
    '    If fgd_composicao_caixa.TextMatrix(i, 0) <> "" And fValidaValor2(fgd_composicao_caixa.TextMatrix(i, 2)) > 0 Then
    '        MovimentoVendaConveniencia.Empresa = g_empresa
    '        MovimentoVendaConveniencia.Data = Format(msk_data.Text, "dd/mm/yyyy")
    '        MovimentoVendaConveniencia.Periodo = Val(cbo_periodo.ItemData(cbo_periodo.ListIndex))
    '        MovimentoVendaConveniencia.NumeroIlha = Val(txt_numero_ilha)
    '        MovimentoVendaConveniencia.TipoMovimento = Val(cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex))
    '        MovimentoVendaConveniencia.CodigoFuncionario = Val(dbcbo_funcionario.BoundText)
    '        MovimentoVendaConveniencia.CodigoComposicao = CLng(fgd_composicao_caixa.TextMatrix(i, 0))
    '        MovimentoVendaConveniencia.valor = CCur(fgd_composicao_caixa.TextMatrix(i, 2))
    '        If MovimentoVendaConveniencia.Incluir Then
    '            lData = msk_data
    '            lPeriodo = cbo_periodo.ItemData(cbo_periodo.ListIndex)
    '            lIlha = Val(txt_numero_ilha)
    '            lTipoMovimento = cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex)
    '            lCodigoComposicao = CLng(fgd_composicao_caixa.TextMatrix(i, 0))
    '        Else
    '            MsgBox "Registro não foi gravado!", vbInformation, "Erro Interno"
    '        End If
    '    End If
    'Next
End Sub
Private Function ExcluiMovimentoCaixa() As Boolean
    Dim xComplemento As String
    Dim xValor As Currency
    
    ExcluiMovimentoCaixa = True
    lNumeroMovimentoCaixa = 0
    xValor = 0
    xComplemento = "VENDA DE LUBRIFICANTES"
    If IntegracaoCaixa.LocalizarNome(g_empresa, xComplemento) Then
        xComplemento = "LUBRIFICANTES Per:" & Val(cbo_periodo.Text) & " Ilha:" & lIlha & " S.Est:" & 2 & " T.Mov:" & 1
        If MovimentoCaixaPista.LocalizarRegistroEspecial(g_empresa, CDate(msk_data.Text), Val(cbo_periodo.Text), 1, xComplemento, IntegracaoCaixa.ContaCredito, "C") Then
            xValor = MovimentoCaixaPista.valor
            lCxDataDigitacao = MovimentoCaixaPista.DataDigitacao
            lCxHoraDigitacao = MovimentoCaixaPista.HoraDigitacao
            lNumeroMovimentoCaixa = MovimentoCaixaPista.NumeroMovimento
            If MovimentoCaixaPista.valor = fValidaValor(fgd_composicao_caixa.TextMatrix(LastRow, 8)) Then
                If Not MovimentoCaixaPista.Excluir(g_empresa, CDate(msk_data.Text), lNumeroMovimentoCaixa) Then
                    MsgBox "Não foi possível excluir o movimento caixa!", vbOKOnly + vbInformation, "Erro de Integridade"
                    ExcluiMovimentoCaixa = False
                End If
            Else
                MovimentoCaixaPista.valor = MovimentoCaixaPista.valor - fValidaValor(fgd_composicao_caixa.TextMatrix(LastRow, 8))
                MovimentoCaixaPista.DataAlteracao = Format(Date, "dd/MM/yyyy")
                MovimentoCaixaPista.HoraAlteracao = Format(Time, "HH:mm:ss")
                If Not MovimentoCaixaPista.Alterar(g_empresa, MovimentoCaixaPista.Data, lNumeroMovimentoCaixa) Then
                    MsgBox "Não foi possível alterar o movimento do caixa!", vbInformation, "Erro de Integridade."
                    ExcluiMovimentoCaixa = False
                End If
            End If
        End If
    Else
        ExcluiMovimentoCaixa = False
    End If
End Function
Private Sub ExibirCelula()
    Static OK As Boolean
    '
    ' Se for celula fixa , sair
    If fgd_composicao_caixa.Col <= fgd_composicao_caixa.FixedCols - 1 Or fgd_composicao_caixa.Row <= fgd_composicao_caixa.FixedRows - 1 Then
       Exit Sub
    End If
    
    If OK Then Exit Sub
    OK = True
    '
    txt_celula.Visible = False
    '
    LastRow = fgd_composicao_caixa.Row
    LastCol = fgd_composicao_caixa.Col
    If LastCol = 0 Then
        txt_celula.MaxLength = 4
    ElseIf LastCol = 2 Then
        txt_celula.MaxLength = 10
    End If
    
    '
    ' Nova Celula
    'With fgd_composicao_caixa
    '  If .TextMatrix(LastRow, 0) = NovaLinha Then
    '    .Rows = .Rows + 1
    '    .TextMatrix(LastRow, 0) = LastRow
    '    .TextMatrix(.Rows - 1, 0) = NovaLinha
    '  End If
    'End With
    '
    Select Case LastCol
        Case Else
        txt_celula.Move fgd_composicao_caixa.CellLeft - Screen.TwipsPerPixelX, fgd_composicao_caixa.CellTop + 1080 - Screen.TwipsPerPixelY, fgd_composicao_caixa.CellWidth + Screen.TwipsPerPixelX * 2, fgd_composicao_caixa.CellHeight + Screen.TwipsPerPixelY * 2
        txt_celula.Text = fgd_composicao_caixa.Text
        'If Len(fgd_composicao_caixa.Text) = 0 Then
        '   If LastRow > 1 Then
        '       txt_celula.Text = fgd_composicao_caixa.TextMatrix(LastRow - 1, LastCol)
        '   End If
        'End If
        txt_celula.Visible = True
        If txt_celula.Visible Then
          txt_celula.ZOrder
          txt_celula.SetFocus
        End If
    End Select
    ControlVisible = True
    OK = False
End Sub
Private Sub Finaliza()
    Set Configuracao = Nothing
    Set Estoque = Nothing
    Set IntegracaoCaixa = Nothing
    Set MovimentoCaixaPista = Nothing
    Set MovimentoLubrificante = Nothing
    Set MovimentoVendaConveniencia = Nothing
    Set MovJustificativa = Nothing
    Set Produto = Nothing
End Sub
Private Sub PreencheCboPeriodo()
    cbo_periodo.Clear
    cbo_periodo.AddItem 1
    cbo_periodo.ItemData(cbo_periodo.NewIndex) = 1
    cbo_periodo.AddItem 2
    cbo_periodo.ItemData(cbo_periodo.NewIndex) = 2
    cbo_periodo.AddItem 3
    cbo_periodo.ItemData(cbo_periodo.NewIndex) = 3
    cbo_periodo.AddItem 4
    cbo_periodo.ItemData(cbo_periodo.NewIndex) = 4
End Sub
Private Sub cbo_periodo_GotFocus()
    SendMessageLong cbo_periodo.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_periodo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        fgd_composicao_caixa.SetFocus
    End If
End Sub
Function IncluiMovimentoCaixa(ByVal pTipoLancamentoPadrao As String) As Boolean
    Dim xComplemento As String
    Dim xContaDebito As String
    Dim xContaCredito As String
    Dim xValor As Currency

    IncluiMovimentoCaixa = False
    xValor = 0
    If IntegracaoCaixa.LocalizarNome(g_empresa, pTipoLancamentoPadrao) Then
        xComplemento = "LUBRIFICANTES Per:" & MovimentoVendaConveniencia.Periodo & " Ilha:" & MovimentoVendaConveniencia.Ilha & " S.Est:" & 2 & " T.Mov:" & 1
        'Caso Exista Deleta e Guarda o Valor
        If MovimentoCaixaPista.LocalizarRegistroEspecial(g_empresa, MovimentoVendaConveniencia.Data, Val(MovimentoVendaConveniencia.Periodo), MovimentoVendaConveniencia.Ilha, xComplemento, IntegracaoCaixa.ContaCredito, "C") Then
            xValor = MovimentoCaixaPista.valor
            If Not MovimentoCaixaPista.Excluir(g_empresa, MovimentoVendaConveniencia.Data, MovimentoCaixaPista.NumeroMovimento) Then
                MsgBox "Não foi possível excluir o movimento caixa!", vbOKOnly + vbInformation, "Erro de Integridade"
            End If
        End If
        xValor = xValor + MovimentoVendaConveniencia.ValorTotal
    Else
        MsgBox "Não existe a integração=" & "VENDA DE LUBRIFICANTES" & ".", vbInformation, "Registro Inexistente"
        Exit Function
    End If
    xComplemento = pTipoLancamentoPadrao
    If IntegracaoCaixa.LocalizarNome(g_empresa, xComplemento) Then
        xContaDebito = IntegracaoCaixa.ContaDebito
        xContaCredito = IntegracaoCaixa.ContaCredito
        MovimentoCaixaPista.valor = xValor
        MovimentoCaixaPista.CodigoHistorico = IntegracaoCaixa.HistoricoPadrao
        MovimentoCaixaPista.DadosInterno = "LUBRI" & "|@|" & 2 & "|@|"
        xComplemento = "LUBRIFICANTES Per:" & MovimentoVendaConveniencia.Periodo & " Ilha:" & MovimentoVendaConveniencia.Ilha & " S.Est:" & 2 & " T.Mov:" & 1
        MovimentoCaixaPista.CodigoLancamentoPadrao = 1
        MovimentoCaixaPista.NumeroDocumento = ""
        MovimentoCaixaPista.Empresa = g_empresa
        MovimentoCaixaPista.Data = MovimentoVendaConveniencia.Data
        MovimentoCaixaPista.NumeroMovimento = 1
        MovimentoCaixaPista.Complemento = Mid(xComplemento, 1, 50)
        MovimentoCaixaPista.NumeroContaDebito = xContaDebito
        MovimentoCaixaPista.NumeroContaCredito = xContaCredito
        MovimentoCaixaPista.CodigoUsuario = g_usuario
        MovimentoCaixaPista.TipoMovimento = 1
        MovimentoCaixaPista.Periodo = MovimentoVendaConveniencia.Periodo
        MovimentoCaixaPista.NumeroIlha = MovimentoVendaConveniencia.Ilha
        MovimentoCaixaPista.DataDigitacao = Format(Now, "dd/mm/yyyy")
        MovimentoCaixaPista.HoraDigitacao = Format(Now, "HH:mm:ss")
        MovimentoCaixaPista.DataAlteracao = "00:00:00"
        MovimentoCaixaPista.HoraAlteracao = "00:00:00"
        If MovimentoCaixaPista.Incluir Then
            'lNumeroMovimentoCaixa = MovimentoCaixaPista.NumeroMovimento
            IncluiMovimentoCaixa = True
        End If
    Else
        MsgBox "Não será possível integrar com o caixa!", vbInformation + vbCritical, "Erro de Integridade"
    End If
End Function
Private Sub LimpaGrid()
    Dim x_sql As String
    Dim i As Integer
    fgd_composicao_caixa.WordWrap = True
    fgd_composicao_caixa.Rows = 2
    fgd_composicao_caixa.Row = 1
    For i = 0 To 9
        fgd_composicao_caixa.Col = i
        fgd_composicao_caixa.Text = ""
    Next
    fgd_composicao_caixa.RowHeight(0) = 500
    fgd_composicao_caixa.Row = 0
    i = 0
    fgd_composicao_caixa.Col = i
    fgd_composicao_caixa.Text = "Data"
    fgd_composicao_caixa.ColWidth(i) = 1000
    fgd_composicao_caixa.ColAlignment(i) = 4
    i = i + 1
    fgd_composicao_caixa.Col = i
    fgd_composicao_caixa.Text = "Número do Cupom"
    fgd_composicao_caixa.ColWidth(i) = 900
    fgd_composicao_caixa.ColAlignment(i) = 7
    i = i + 1
    fgd_composicao_caixa.Col = i
    fgd_composicao_caixa.Text = "Ordem"
    fgd_composicao_caixa.ColWidth(i) = 600
    fgd_composicao_caixa.ColAlignment(i) = 7
    i = i + 1
    fgd_composicao_caixa.Col = i
    fgd_composicao_caixa.Text = "Período"
    fgd_composicao_caixa.ColWidth(i) = 700
    fgd_composicao_caixa.ColAlignment(i) = 4
    i = i + 1
    fgd_composicao_caixa.Col = i
    fgd_composicao_caixa.Text = "Código do Produto"
    fgd_composicao_caixa.ColWidth(i) = 900
    fgd_composicao_caixa.ColAlignment(i) = 7
    i = i + 1
    fgd_composicao_caixa.Col = i
    fgd_composicao_caixa.Text = "Nome do Produto"
    fgd_composicao_caixa.ColWidth(i) = 2900
    fgd_composicao_caixa.ColAlignment(i) = 1
    i = i + 1
    fgd_composicao_caixa.Col = i
    fgd_composicao_caixa.Text = "Valor Unitário"
    fgd_composicao_caixa.ColWidth(i) = 900
    fgd_composicao_caixa.ColAlignment(i) = 7
    i = i + 1
    fgd_composicao_caixa.Col = i
    fgd_composicao_caixa.Text = "Quantidade"
    fgd_composicao_caixa.ColWidth(i) = 900
    fgd_composicao_caixa.ColAlignment(i) = 7
    i = i + 1
    fgd_composicao_caixa.Col = i
    fgd_composicao_caixa.Text = "Valor Total"
    fgd_composicao_caixa.ColWidth(i) = 900
    fgd_composicao_caixa.ColAlignment(i) = 7
    i = i + 1
    fgd_composicao_caixa.Col = i
    fgd_composicao_caixa.Text = "Situação"
    fgd_composicao_caixa.ColWidth(i) = 900
    fgd_composicao_caixa.ColAlignment(i) = 4
    'x'lbl_total_nota.Caption = ""
    txt_celula.Visible = False
    fgd_composicao_caixa.Row = 1
    fgd_composicao_caixa.Col = 0
    fgd_composicao_caixa.Text = ""
End Sub
Private Sub cbo_periodo_LostFocus()
    If VerificaLiberacaoDigitacao2 Then
        AtualizaGrid
    End If
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    msk_data.Text = RetiraGString(1)
    cbo_periodo.SetFocus
    g_string = ""
End Sub
Function VerificaLiberacaoDigitacao2() As Boolean
    VerificaLiberacaoDigitacao2 = False
    If g_nivel_acesso <= 4 Then
        VerificaLiberacaoDigitacao2 = True
        Exit Function
    End If
    If msk_data.Text < g_cfg_data_i Or msk_data.Text > g_cfg_data_f Then
        MsgBox "A data do movimento deve estar entre " & Format(g_cfg_data_i, "dd/mm/yyyy") & " a " & Format(g_cfg_data_f, "dd/mm/yyyy") & ".", 64, "Digitação Não Autorizada!"
        msk_data.SetFocus
    ElseIf cbo_periodo.Text < g_cfg_periodo_i Or cbo_periodo.Text > g_cfg_periodo_f Then
        MsgBox "O período deve estar entre " & g_cfg_periodo_i & " ao " & g_cfg_periodo_f & ".", 64, "Digitação Não Autorizada!"
        cbo_periodo.SetFocus
    Else
        VerificaLiberacaoDigitacao2 = True
    End If
End Function
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub ProximaCelula()
    If fgd_composicao_caixa.Col < fgd_composicao_caixa.Cols - 2 Then
        fgd_composicao_caixa.Col = LastCol + 1
    Else
        fgd_composicao_caixa.Col = 9
        If fgd_composicao_caixa.Row >= fgd_composicao_caixa.Rows - 1 Then
            fgd_composicao_caixa.Row = fgd_composicao_caixa.Row - 1
        End If
        fgd_composicao_caixa.Row = fgd_composicao_caixa.Row + 1
    End If
    fgd_composicao_caixa.SetFocus
End Sub
Private Sub TotalizaGrid()
    Dim x_total As Currency
    Dim i As Integer
    x_total = 0
    With fgd_composicao_caixa
        For i = 1 To (.Rows - 1)
            If Len(.TextMatrix(i, 0)) > 0 Then
                If .TextMatrix(i, 9) = "Normal" Then
                    x_total = x_total + fValidaValor(.TextMatrix(i, 8))
                End If
            End If
        Next
    End With
    lbl_total.Caption = Format(x_total, "###,###,##0.00")
    Call GravaAuditoria(1, Me.name, 5, "Total:" & Format(x_total, "###,###,##0.00") & " Data:" & msk_data.Text & " Per:" & cbo_periodo.ItemData(cbo_periodo.ListIndex))
End Sub
Private Sub fgd_composicao_caixa_Click()
    ' Quando clicar uma vez
    ' atribui o valor selecionado
    lMarcaCelula = True
    If fgd_composicao_caixa.Col = 9 Then
        LastRow = fgd_composicao_caixa.Row
        LastCol = fgd_composicao_caixa.Col
        txt_celula.Visible = False
    End If
    'AtribuiValorCelula
End Sub
Private Sub fgd_composicao_caixa_DblClick()
    'editar ao clicar duas vezes
    lMarcaCelula = True
    If fgd_composicao_caixa.Col = 9 Then
        '8 - Cancelado/Normal
        LastRow = fgd_composicao_caixa.Row
        LastCol = fgd_composicao_caixa.Col
        txt_celula.Visible = False
        ExibirCelula
    End If
End Sub
Private Sub fgd_composicao_caixa_KeyPress(KeyAscii As Integer)
    lMarcaCelula = True
    Select Case KeyAscii
    ' Editar ao teclar ENTER
    Case vbKeyReturn
        KeyAscii = 0
        If fgd_composicao_caixa.Col = 9 Then
            ExibirCelula
        End If
    ' Cancelar ao pressionar ESC
    Case vbKeyEscape
        KeyAscii = 0
        AtribuiValorCelula
    ' Editar ao pressinar qualquer tecla
    Case 32 To 255
        lMarcaCelula = False
        If fgd_composicao_caixa.Col = 9 Then
            ExibirCelula
            With txt_celula
                If .Visible Then
                    .Text = Chr$(KeyAscii)
                    .SelStart = Len(.Text) + 1
                End If
            End With
        End If
    End Select
End Sub
Private Sub fgd_composicao_caixa_Scroll()
    ' Ver se a coluna esta visivel
    ' entao ocultar os controles
    '
    If fgd_composicao_caixa.ColIsVisible(LastCol) = False Then
        txt_celula.Visible = False
        Exit Sub
    End If
    If fgd_composicao_caixa.RowIsVisible(LastRow) = False Then
        txt_celula.Visible = False
        Exit Sub
    End If
    ' ver se estava visivel antes de ocultar
    ' e posicionar na mesma celula
    If ControlVisible Then
        ExibirCelula
    End If
End Sub
Private Sub Form_Activate()
    If g_empresa <> lEmpresa Then
        lFlagMovimento = 0
    End If
    If lFlagMovimento = 0 Then
        AtualizaConstantes
        lOpcao = 0
        lEmpresa = g_empresa
        msk_data.Text = RetiraGString(1)
        cbo_periodo.ListIndex = RetiraGString(2) - 1
        lCodigoFuncionario = Val(RetiraGString(3))
        lIlha = Val(RetiraGString(4))
        lOrigemVenda = RetiraGString(5)
        'If MovimentoVendaConveniencia.LocalizarUltimo(g_empresa) Then
        '    AtualTela
            AtualizaGrid
        '    AtivaBotoes
        'Else
        '    cmd_novo.Enabled = True
        '    cmd_sair.Enabled = True
        'End If
        'If cmd_novo.Enabled Then
        '    cmd_novo.SetFocus
        'End If
    Else
        lFlagMovimento = 0
    End If
    Screen.MousePointer = 1
End Sub
Private Sub Form_Deactivate()
    lFlagMovimento = 1
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF6 And lOpcao = 0 Then
        KeyCode = 0
        cmd_sair_Click
    End If
End Sub
Private Sub Form_Load()
    CentraForm Me
    
    PreencheCboPeriodo
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_GotFocus()
    msk_data.SelStart = 0
    msk_data.SelLength = 5
End Sub
Private Sub msk_data_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_periodo.SetFocus
    End If
End Sub
Private Sub txt_celula_GotFocus()
    With txt_celula
        If LastCol = 9 Then
            .MaxLength = 9
        End If
        If lMarcaCelula Then
            If .Text = "Cancelado" Then
                .Text = "X"
            Else
                .Text = " "
            End If
            .SelStart = 0
            .SelLength = Len(.Text)
        End If
    End With
End Sub
Private Sub txt_celula_KeyPress(KeyAscii As Integer)
    ' ao pressionar ENTER aceitar a entrada de dados
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        fgd_composicao_caixa.SetFocus
    ' ESC, cancela a edição
    ElseIf KeyAscii = vbKeyEscape Then
        KeyAscii = 0
        txt_celula.Visible = False
        ControlVisible = False
    End If
    'If LastCol = 0 Then
    '    Call ValidaInteiro(KeyAscii)
    'ElseIf LastCol = 2 Then
    '    If KeyAscii = 46 Then
    '        KeyAscii = 44
    '    End If
    '    Call ValidaValor(KeyAscii)
    'End If
End Sub
Private Sub txt_celula_LostFocus()
    'Cancelado/Normal
    If LastCol = 9 Then
        If UCase(txt_celula.Text) = "X" Then
            txt_celula.Text = "Cancelado"
            g_string = "Cancela Venda de Conveniência" & "|@|"
            g_string = g_string & Me.name & "|@|" & lCodigoFuncionario & "|@|"
            MovimentoJustificativa.Show 1
            If RetiraGString(1) = "OK" Then
                Call CriaLogConveniencia(Time & " - Cancela: Produto:" & fgd_composicao_caixa.TextMatrix(LastRow, 5) & " - Vlr Unit:" & fgd_composicao_caixa.TextMatrix(LastRow, 6) & " - Qtd:" & fgd_composicao_caixa.TextMatrix(LastRow, 7) & " - Vlr Total:" & fgd_composicao_caixa.TextMatrix(LastRow, 8), "", "")
                Call GravaAuditoria(1, Me.name, 10, "Venda Cancelada: Data:" & fgd_composicao_caixa.TextMatrix(LastRow, 0) & " Per:" & fgd_composicao_caixa.TextMatrix(LastRow, 3))
                Call GravaAuditoria(1, Me.name, 10, "N.Cupom:" & fgd_composicao_caixa.TextMatrix(LastRow, 1) & " Prod:" & fgd_composicao_caixa.TextMatrix(LastRow, 4) & " Qtd:" & fgd_composicao_caixa.TextMatrix(LastRow, 7) & " Tot:" & fgd_composicao_caixa.TextMatrix(LastRow, 8))
                lNumeroJustificativa = CLng(RetiraGString(2))
                g_string = ""
                AtribuiValorCelula
                AlteraMovimentoVendaConveniencia (True)
                AtualizaTabelaVendaProduto (True)
            Else
                txt_celula.Text = "Normal"
                AtribuiValorCelula
            End If
        Else
            Call GravaAuditoria(1, Me.name, 10, "Venda retornada: Data:" & fgd_composicao_caixa.TextMatrix(LastRow, 0) & " Per:" & fgd_composicao_caixa.TextMatrix(LastRow, 3))
            Call GravaAuditoria(1, Me.name, 10, "N.Cupom:" & fgd_composicao_caixa.TextMatrix(LastRow, 1) & " Prod:" & fgd_composicao_caixa.TextMatrix(LastRow, 4) & " Qtd:" & fgd_composicao_caixa.TextMatrix(LastRow, 7) & " Tot:" & fgd_composicao_caixa.TextMatrix(LastRow, 8))
            lNumeroJustificativa = 0
            txt_celula.Text = "Normal"
            AtribuiValorCelula
            AlteraMovimentoVendaConveniencia (False)
            AtualizaTabelaVendaProduto (False)
        End If
        TotalizaGrid
    End If
    If LastCol <> 1 Then
        ProximaCelula
    End If
End Sub
