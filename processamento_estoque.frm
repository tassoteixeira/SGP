VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form processamento_estoque 
   Caption         =   "Processamento de Estoque"
   ClientHeight    =   5805
   ClientLeft      =   165
   ClientTop       =   285
   ClientWidth     =   7410
   Icon            =   "processamento_estoque.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   Picture         =   "processamento_estoque.frx":030A
   ScaleHeight     =   5805
   ScaleWidth      =   7410
   Begin VB.Frame frmDados 
      Height          =   4755
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   7275
      Begin VB.CheckBox chkSomaEntradaInventario 
         Caption         =   "&Soma as &Entradas (Inventário do Estoque Inicial do Dia) no Estoque Atual"
         Height          =   300
         Left            =   180
         TabIndex        =   2
         Top             =   600
         Width           =   5595
      End
      Begin VB.CheckBox chk_calcula_entrada 
         Caption         =   "&Calcula as Entradas do Período"
         Height          =   300
         Left            =   180
         TabIndex        =   4
         Top             =   1320
         Width           =   5595
      End
      Begin VB.CheckBox chk_calcula_saida 
         Caption         =   "Ca&lcula as Saidas do Período"
         Height          =   300
         Left            =   180
         TabIndex        =   5
         Top             =   1680
         Width           =   5595
      End
      Begin VB.CheckBox chkVendaInterna 
         Caption         =   "Movimentação Interna"
         Height          =   300
         Left            =   480
         TabIndex        =   6
         Top             =   1980
         Width           =   3675
      End
      Begin VB.CheckBox chkVendaCupom 
         Caption         =   "Movimentação de Cupom Fiscal / NFe"
         Height          =   300
         Left            =   480
         TabIndex        =   7
         Top             =   2280
         Width           =   3675
      End
      Begin VB.CheckBox chkVendaConveniencia 
         Caption         =   "Movimentação de Venda de Conveniencia"
         Height          =   300
         Left            =   480
         TabIndex        =   8
         Top             =   2580
         Width           =   3735
      End
      Begin VB.CheckBox chkMoveInventarioContabil 
         Caption         =   "&Move Inventário Contábil para o Estoque Atual"
         Height          =   300
         Left            =   180
         TabIndex        =   3
         Top             =   960
         Width           =   5595
      End
      Begin VB.CheckBox chkTransferenciaSubEstoque 
         Caption         =   "Calcula &Transferência entre Sub-Estoque"
         Height          =   300
         Left            =   180
         TabIndex        =   9
         Top             =   2880
         Width           =   5595
      End
      Begin VB.ComboBox cboProduto 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   4320
         Width           =   5475
      End
      Begin VB.ComboBox cboGrupo 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   3900
         Width           =   5475
      End
      Begin VB.CheckBox chkCalculoInverso 
         Caption         =   "Cálculo Inverso"
         Height          =   300
         Left            =   4260
         TabIndex        =   15
         Top             =   3480
         Width           =   2775
      End
      Begin VB.CheckBox chk_zera_estoque 
         Caption         =   "&Zera o Estoque Atual"
         Height          =   300
         Left            =   180
         TabIndex        =   1
         Top             =   240
         Width           =   5595
      End
      Begin MSMask.MaskEdBox msk_data_inicial 
         Height          =   300
         Left            =   180
         TabIndex        =   12
         Top             =   3480
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
      Begin MSMask.MaskEdBox msk_data_final 
         Height          =   300
         Left            =   2400
         TabIndex        =   14
         Top             =   3480
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
         Caption         =   "&Data inicial"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   11
         Top             =   3270
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "D&ata final"
         Height          =   195
         Index           =   8
         Left            =   2400
         TabIndex        =   13
         Top             =   3270
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "P&roduto"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   18
         Top             =   4320
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "&Grupo"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   3900
         Width           =   1515
      End
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   1680
      Picture         =   "processamento_estoque.frx":0750
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Confirma o processamento."
      Top             =   4860
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4020
      Picture         =   "processamento_estoque.frx":1D5A
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   4860
      Width           =   795
   End
End
Attribute VB_Name = "processamento_estoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Estoque As New cEstoque
Private SubEstoque As New cSubEstoque
Private Cfop As New cCfop
Dim lSQL As String
Dim rs As New adodb.Recordset
Dim lNomeOperacao As String

Private Function CalculaEstoqueSubEstoque(ByVal pCalcEstoque As Boolean, ByVal pCalcSubEstoque As Boolean, ByVal pCodigoProduto As Long, ByVal pCodigoTipoSubEstoque As Integer, ByVal pQuantidade As Currency, ByVal pSoma As Boolean) As Boolean
    CalculaEstoqueSubEstoque = False
    If pSoma Then
        lNomeOperacao = "Adicionar"
    Else
        lNomeOperacao = "Subtrair"
    End If
    
    If pCodigoProduto = 0 Then
        CalculaEstoqueSubEstoque = True
        Exit Function
    End If
    
    'Calcula Estoque
    If pCalcEstoque Then
        If Estoque.AlterarQuantidade(g_empresa, pCodigoProduto, pQuantidade, pSoma) Then
            CalculaEstoqueSubEstoque = True
        Else
            CalculaEstoqueSubEstoque = False
            Call GravaAuditoria(1, Me.name, 26, "Erro " & lNomeOperacao & " Estoque. Cod:" & pCodigoProduto & " Qtd:" & pQuantidade)
            MsgBox "Não foi possível " & lNomeOperacao & " Estoque o produto:" & pCodigoProduto, vbInformation, "Erro de Integridade!"
            Exit Function
        End If
    End If
                
    'Calcula SubEstoque
    If pCalcSubEstoque Then
        If SubEstoque.AlterarQuantidade(g_empresa, pCodigoProduto, pCodigoTipoSubEstoque, pQuantidade, pSoma) Then
            CalculaEstoqueSubEstoque = True
        Else
            CalculaEstoqueSubEstoque = False
            Call GravaAuditoria(1, Me.name, 26, "Erro " & lNomeOperacao & " SubEstoque. Cod:" & pCodigoProduto & " SubEst:" & pCodigoTipoSubEstoque & " Qtd:" & pQuantidade)
            MsgBox "Não foi possível " & lNomeOperacao & " SubEstoque!", vbInformation, "Erro de Integridade!"
            Exit Function
        End If
    End If
End Function
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set Estoque = Nothing
    Set SubEstoque = Nothing
    Set Cfop = Nothing
End Sub
Private Sub Processamento()
    If chk_zera_estoque.Value = 1 Then
        ProcessamentoZeraEstoque
    End If
    If chkSomaEntradaInventario.Value = 1 Then
        ProcessamentoSomaEntradaInventario
    End If
    If chkMoveInventarioContabil = 1 Then
        ProcessamentoMoveInventarioContabil
    End If
    If chk_calcula_entrada.Value = 1 Then
        ProcessamentoCalculaEntradaEstoque
    End If
    If chk_calcula_saida.Value = 1 Then
        If chkVendaInterna.Value = 1 Then
            ProcessamentoCalculaSaidaInterna
        End If
        If chkVendaCupom.Value = 1 Then
            ProcessamentoCalculaSaidaCupom
            ProcessamentoCalculaNFe
        End If
        If chkVendaConveniencia.Value = 1 Then
            ProcessamentoCalculaSaidaConveniencia
        End If
    End If
    If chkTransferenciaSubEstoque.Value = 1 Then
        ProcessamentoTransfSubEstoque
    End If
End Sub
Private Sub ProcessamentoCalculaEntradaEstoque()
Dim rsEntradaProduto As adodb.Recordset
Dim xSQL As String
Dim xAtualizou As Boolean
Dim xOperacaoSoma As Boolean

On Error GoTo trata_erro
    
    If (MsgBox("Na empresa " & g_nome_empresa & Chr(10) & "Será calculado as entradas de " & msk_data_inicial & " até " & msk_data_final & " para todo estoque." & Chr(10) & Chr(10) & "Deseja realmente fazer este processamento?", vbYesNo + 256, "Calcula Entradas para o Estoque!")) = 6 Then
        Call GravaAuditoria(1, Me.name, 26, "Calcula Entradas Empresa:" & g_empresa & "-" & Mid(g_nome_empresa, 1, 20) & " De:" & msk_data_inicial.Text & " a " & msk_data_final.Text)
        xAtualizou = True
        
        Conectar.IniciaTransacao
        Set rsEntradaProduto = New adodb.Recordset
        xSQL = ""
        'xSQL = xSQL & "SELECT [Codigo do Produto], Quantidade, SubEstoque" '..old 21/07/2015
        xSQL = xSQL & "SELECT [Codigo do Produto], SubEstoque, CFOP, SUM(Quantidade) AS Quantidade" 'new 21/07/2015
        xSQL = xSQL & "  FROM Entrada_Produto"
        xSQL = xSQL & " WHERE Empresa = " & g_empresa
        xSQL = xSQL & "   AND [Data da Entrada] >= " & preparaData(CDate(msk_data_inicial.Text))
        xSQL = xSQL & "   AND [Data da Entrada] <= " & preparaData(CDate(msk_data_final.Text))
        'xSQL = xSQL & "   AND [Tipo da Entrada] IN ('1', '2', '4')" 'old 21/07/2015
        xSQL = xSQL & "   AND [Tipo da Entrada] = " & preparaTexto("1") 'new 21/07/2015
        
        If cboProduto.ItemData(cboProduto.ListIndex) > 0 Then 'new 20/07/2015 daqui...
            xSQL = xSQL & "    AND [Codigo do Produto] = " & cboProduto.ItemData(cboProduto.ListIndex)
        End If
        If cboGrupo.ItemData(cboGrupo.ListIndex) > 0 Then
            xSQL = xSQL & " AND [Codigo do Produto] IN"
            xSQL = xSQL & "   ("
            xSQL = xSQL & "    SELECT Produto.Codigo"
            xSQL = xSQL & "      FROM Produto"
            xSQL = xSQL & "     WHERE Produto.[Codigo do Grupo] = " & cboGrupo.ItemData(cboGrupo.ListIndex)
            xSQL = xSQL & "   )"
        End If '... ate aqui new 20/07/2015
        'xSQL = xSQL & " ORDER BY [Codigo do Produto], [Data da Entrada]" ..old 21/07/2015
        xSQL = xSQL & " GROUP BY [Codigo do Produto], SubEstoque, CFOP" 'new 21/07/2015
        xSQL = xSQL & " ORDER BY [Codigo do Produto], SubEstoque, CFOP" 'new 21/07/2015

        Set rsEntradaProduto = Conectar.RsConexao(xSQL)
        If rsEntradaProduto.RecordCount > 0 Then
            Do Until rsEntradaProduto.EOF
            
                xOperacaoSoma = True
                If Cfop.LocalizarCodigo(rsEntradaProduto("CFOP").Value) Then
                    If Cfop.Operacao = "-" Then
                        xOperacaoSoma = False
                    ElseIf Cfop.Operacao = "+" Then
                        xOperacaoSoma = True
                    End If
                Else
                    MsgBox "CFOP inexistente!" + vbCrLf + "Cfop: " + rsEntradaProduto("CFOP").Value, vbCritical, "Erro de Integridade"
                End If
            
                If Cfop.Operacao <> "=" Then
                    If chkCalculoInverso.Value = 1 Then
                        xOperacaoSoma = Not xOperacaoSoma
                    End If
                    If CalculaEstoqueSubEstoque(True, True, rsEntradaProduto("Codigo do Produto").Value, rsEntradaProduto("SubEstoque").Value, rsEntradaProduto("Quantidade").Value, xOperacaoSoma) Then
                    Else
    'aquiaqui                    xAtualizou = False
    '                    Conectar.CancelaTransacao
    '                    MsgBox "Todo o processamento de cálculo de entrada está cancelado!", vbCritical, "Erro de Integridade!"
    '                    Call GravaAuditoria(1, Me.name, 26, "Todo o processamento de cálculo de entrada está cancelado!")
    '                    Exit Do
                    End If
                End If
'                If Estoque.LocalizarCodigo(g_empresa, rsEntradaProduto("Codigo do Produto").Value) Then
'                    Estoque.Quantidade = Estoque.Quantidade + rsEntradaProduto("Quantidade").Value
'                    If Estoque.Alterar(g_empresa, rsEntradaProduto("Codigo do Produto").Value) Then
'                        If SubEstoque.LocalizarCodigo(g_empresa, rsEntradaProduto("Codigo do Produto").Value, rsEntradaProduto("SubEstoque").Value) Then
'                            SubEstoque.Quantidade = SubEstoque.Quantidade + rsEntradaProduto("Quantidade").Value
'                            If Not SubEstoque.Alterar(g_empresa, rsEntradaProduto("Codigo do Produto").Value, rsEntradaProduto("SubEstoque").Value) Then
'                                xAtualizou = False
'                                Conectar.CancelaTransacao
'                                MsgBox "Não foi possível alterar o Sub-Estoque!", vbInformation, "Erro de Integridade!"
'                                Exit Do
'                            End If
'                        End If
'                    Else
'                        xAtualizou = False
'                        Conectar.CancelaTransacao
'                        MsgBox "Não foi possível alterar o estoque!", vbInformation, "Erro de Integridade!"
'                        Exit Do
'                    End If
'                End If
                rsEntradaProduto.MoveNext
            Loop
        End If
        If xAtualizou Then
            Conectar.ConfirmaTransacao
        End If
        rsEntradaProduto.Close
        Set rsEntradaProduto = Nothing
        MsgBox "A empresa " & g_nome_empresa & Chr(10) & "Está com as entradas calculadas para o estoque.", vbInformation, "Operação Concluida!"
    End If
    Exit Sub

trata_erro:
    Conectar.CancelaTransacao
    MsgBox Err.Number & " - " & Err.Description
End Sub
Private Sub ProcessamentoCalculaNFe()
Dim rsNFeItem As adodb.Recordset
Dim xSQL As String
Dim xAtualizou As Boolean
Dim xOperacaoSoma As Boolean

On Error GoTo trata_erro
    
    If (MsgBox("Na empresa " & g_nome_empresa & Chr(10) & "Será calculado as NFe de " & msk_data_inicial.Text & " até " & msk_data_final.Text & " para todo estoque." & Chr(10) & Chr(10) & "Deseja realmente fazer este processamento?", vbYesNo + 256, "Calcula NFe para o Estoque!")) = vbYes Then
        Call GravaAuditoria(1, Me.name, 26, "NFe Empresa:" & g_empresa & "-" & Mid(g_nome_empresa, 1, 20) & " De:" & msk_data_inicial.Text & " a " & msk_data_final.Text)
        
        Set rsNFeItem = New adodb.Recordset
        xAtualizou = True
        Conectar.IniciaTransacao
        xSQL = ""
        xSQL = xSQL & "SELECT [Codigo de Produto],  SubEstoque, CFOP, SUM (Quantidade) As Quantidade" 'new 21/07/2015
        xSQL = xSQL & "  FROM MovimentoNotaFiscalSaidaItem"
        xSQL = xSQL & " WHERE Empresa = " & g_empresa
        xSQL = xSQL & "   AND Data >= " & preparaData(CDate(msk_data_inicial.Text))
        xSQL = xSQL & "   AND Data <= " & preparaData(CDate(msk_data_final.Text))
        xSQL = xSQL & "   AND [Tipo de Combustivel] = " & preparaTexto("")
        xSQL = xSQL & "   AND Cancelado = " & preparaBooleano(False)
        xSQL = xSQL & "   AND [CFOP] NOT IN ( " & preparaTexto("5929")
        xSQL = xSQL & ", " & preparaTexto("6929")
        xSQL = xSQL & " )"
        If cboProduto.ItemData(cboProduto.ListIndex) > 0 Then 'new 21/07/2015 daqui ...
            xSQL = xSQL & "    AND [Codigo de Produto] = " & cboProduto.ItemData(cboProduto.ListIndex)
        Else
            If cboGrupo.ItemData(cboGrupo.ListIndex) > 0 Then
                xSQL = xSQL & " AND [Codigo de Produto] IN"
                xSQL = xSQL & "   ("
                xSQL = xSQL & "    SELECT Produto.Codigo"
                xSQL = xSQL & "      FROM Produto"
                xSQL = xSQL & "     WHERE Produto.[Codigo do Grupo] = " & cboGrupo.ItemData(cboGrupo.ListIndex)
                xSQL = xSQL & "   )"
            End If
        End If
        xSQL = xSQL & " GROUP BY [Codigo de Produto], SubEstoque, CFOP" 'new 21/07/2015
        xSQL = xSQL & " ORDER BY [Codigo de Produto], SubEstoque, CFOP" 'new 21/07/2015
        Set rsNFeItem = Conectar.RsConexao(xSQL)
        If rsNFeItem.RecordCount > 0 Then
            Do Until rsNFeItem.EOF
                xOperacaoSoma = False
                If Cfop.LocalizarCodigo(rsNFeItem("CFOP").Value) Then
                    If Cfop.Operacao = "-" Then
                        xOperacaoSoma = False
                    ElseIf Cfop.Operacao = "+" Then
                        xOperacaoSoma = True
                    End If
                Else
                    MsgBox "CFOP inexistente!" + vbCrLf + "Cfop: " + rsNFeItem("CFOP").Value, vbCritical, "Erro de Integridade"
                End If
            
                If Cfop.Operacao <> "=" Then
                    If chkCalculoInverso.Value = 1 Then
                        xOperacaoSoma = Not xOperacaoSoma
                    End If
                    If CalculaEstoqueSubEstoque(True, True, rsNFeItem("Codigo de Produto").Value, 1, rsNFeItem("Quantidade").Value, xOperacaoSoma) Then
                    Else
                        xAtualizou = False
                        Conectar.CancelaTransacao
                        MsgBox "Todo o processamento de cálculo da NFe está cancelado!", vbCritical, "Erro de Integridade!"
                        Call GravaAuditoria(1, Me.name, 26, "Todo o processamento de cálculo da NFe está cancelado!")
                        Exit Do
                    End If
                End If
                rsNFeItem.MoveNext
            Loop
        End If
        If xAtualizou Then
            Conectar.ConfirmaTransacao
        End If
        rsNFeItem.Close
        Set rsNFeItem = Nothing
        MsgBox "A empresa " & g_nome_empresa & Chr(10) & "Está com as NFe calculadas para o estoque.", vbInformation, "Operação Concluida!"
    End If
    Exit Sub

trata_erro:
    Conectar.CancelaTransacao
    MsgBox Err.Number & " - " & Err.Description
End Sub
Private Sub ProcessamentoCalculaSaidaCupom()
Dim rsMovimentoCupomFiscal As adodb.Recordset
Dim xSQL As String
Dim xAtualizou As Boolean
Dim xOperacaoSoma As Boolean

On Error GoTo trata_erro
    
    If (MsgBox("Na empresa " & g_nome_empresa & Chr(10) & "Será calculado as saidas de cupom de " & msk_data_inicial.Text & " até " & msk_data_final.Text & " para todo estoque." & Chr(10) & Chr(10) & "Deseja realmente fazer este processamento?", vbYesNo + 256, "Calcula Saidas de Cupom para o Estoque!")) = 6 Then
        Call GravaAuditoria(1, Me.name, 26, "Saídas Cupom Empresa:" & g_empresa & "-" & Mid(g_nome_empresa, 1, 20) & " De:" & msk_data_inicial.Text & " a " & msk_data_final.Text)
        xOperacaoSoma = False
        If chkCalculoInverso.Value = 1 Then
            xOperacaoSoma = True
        End If
        
        Set rsMovimentoCupomFiscal = New adodb.Recordset
        xAtualizou = True
        Conectar.IniciaTransacao
        xSQL = ""
        xSQL = xSQL & "SELECT [Codigo do Produto], [Tipo do SubEstoque] , SUM (Quantidade) As Quantidade" 'new 21/07/2015
        xSQL = xSQL & "  FROM Movimento_Cupom_Fiscal"
        xSQL = xSQL & " WHERE Empresa = " & g_empresa
        xSQL = xSQL & "   AND Data >= " & preparaData(CDate(msk_data_inicial.Text))
        xSQL = xSQL & "   AND Data <= " & preparaData(CDate(msk_data_final.Text))
        xSQL = xSQL & "   AND [Cupom Cancelado] = " & preparaBooleano(False)
        xSQL = xSQL & "   AND [Item Cancelado] = " & preparaBooleano(False)
        If cboProduto.ItemData(cboProduto.ListIndex) > 0 Then 'new 21/07/2015 daqui ...
            xSQL = xSQL & "    AND [Codigo do Produto] = " & cboProduto.ItemData(cboProduto.ListIndex)
        Else
            If cboGrupo.ItemData(cboGrupo.ListIndex) > 0 Then
                xSQL = xSQL & "    AND [Codigo do Grupo] = " & cboGrupo.ItemData(cboGrupo.ListIndex)
            End If '... ate aqui new 21/07/2015
        End If
        xSQL = xSQL & " GROUP BY [Codigo do Produto], [Tipo do SubEstoque]" 'new 21/07/2015
        xSQL = xSQL & " ORDER BY [Codigo do Produto], [Tipo do SubEstoque]" 'new 21/07/2015
        'aquiaquiaqui
        Set rsMovimentoCupomFiscal = Conectar.RsConexaoTimeOut(xSQL, 300)
        If rsMovimentoCupomFiscal.RecordCount > 0 Then
            Do Until rsMovimentoCupomFiscal.EOF
                If CalculaEstoqueSubEstoque(True, True, rsMovimentoCupomFiscal("Codigo do Produto").Value, 1, rsMovimentoCupomFiscal("Quantidade").Value, xOperacaoSoma) Then
                Else
                    xAtualizou = False
                    Conectar.CancelaTransacao
                    MsgBox "Todo o processamento de cálculo da saída ECF está cancelado!", vbCritical, "Erro de Integridade!"
                    Call GravaAuditoria(1, Me.name, 26, "Todo o processamento de cálculo da saída ECF está cancelado!")
                    Exit Do
                End If
'                If Estoque.AlterarQuantidade(g_empresa, rsMovimentoCupomFiscal("Codigo do Produto").Value, rsMovimentoCupomFiscal("Quantidade").Value, False) Then
'                    'If Not SubEstoque.AlterarQuantidade(g_empresa, rsMovimentoCupomFiscal("Codigo do Produto").Value, rsMovimentoCupomFiscal("Tipo do SubEstoque").Value, rsMovimentoCupomFiscal("Quantidade").Value, False) Then
'                    If Not SubEstoque.AlterarQuantidade(g_empresa, rsMovimentoCupomFiscal("Codigo do Produto").Value, 1, rsMovimentoCupomFiscal("Quantidade").Value, False) Then
'                        xAtualizou = False
'                        Conectar.CancelaTransacao
'                        MsgBox "Não foi possível alterar o Sub-Estoque!", vbInformation, "Erro de Integridade!"
'                        Exit Do
'                    End If
'                Else
'                    xAtualizou = False
'                    Conectar.CancelaTransacao
'                    MsgBox "Não foi possível alterar o estoque!", vbInformation, "Erro de Integridade!"
'                    Exit Do
'                End If
                rsMovimentoCupomFiscal.MoveNext
            Loop
        End If
        If xAtualizou Then
            Conectar.ConfirmaTransacao
        End If
        rsMovimentoCupomFiscal.Close
        Set rsMovimentoCupomFiscal = Nothing
        MsgBox "A empresa " & g_nome_empresa & Chr(10) & "Está com as saidas de cupom calculadas para o estoque.", vbInformation, "Operação Concluida!"
    End If
    Exit Sub

trata_erro:
    Conectar.CancelaTransacao
    MsgBox Err.Number & " - " & Err.Description
End Sub
Private Sub ProcessamentoCalculaSaidaConveniencia()
Dim rsMovimentoVendaConveniencia As adodb.Recordset
Dim xSQL As String
Dim xAtualizou As Boolean
Dim xOperacaoSoma As Boolean

On Error GoTo trata_erro
    
    If (MsgBox("Na empresa " & g_nome_empresa & Chr(10) & "Será calculado as saidas de conveniencia de " & msk_data_inicial.Text & " até " & msk_data_final.Text & " para todo estoque." & Chr(10) & Chr(10) & "Deseja realmente fazer este processamento?", vbYesNo + 256, "Calcula Saidas de Conveniencia para o Estoque!")) = 6 Then
        Call GravaAuditoria(1, Me.name, 26, "Saídas Conveniência Empresa:" & g_empresa & "-" & Mid(g_nome_empresa, 1, 20) & " De:" & msk_data_inicial.Text & " a " & msk_data_final.Text)
        xOperacaoSoma = False
        If chkCalculoInverso.Value = 1 Then
            xOperacaoSoma = True
        End If
        
        Set rsMovimentoVendaConveniencia = New adodb.Recordset
        xAtualizou = True
        Conectar.IniciaTransacao
        xSQL = ""
        'xSQL = xSQL & "SELECT [Codigo do Produto], Quantidade"
        xSQL = xSQL & "SELECT [Codigo do Produto], SUM (Quantidade) As Quantidade" 'new 21/07/2015
        xSQL = xSQL & "  FROM Movimento_Venda_Conveniencia"
        xSQL = xSQL & " WHERE Empresa = " & g_empresa
        xSQL = xSQL & "   AND Data >= " & preparaData(CDate(msk_data_inicial.Text))
        xSQL = xSQL & "   AND Data <= " & preparaData(CDate(msk_data_final.Text))
        xSQL = xSQL & "   AND [Cupom Cancelado] = " & preparaBooleano(False)
        xSQL = xSQL & "   AND [Item Cancelado] = " & preparaBooleano(False)
        If cboProduto.ItemData(cboProduto.ListIndex) > 0 Then
            xSQL = xSQL & "    AND [Codigo do Produto] = " & cboProduto.ItemData(cboProduto.ListIndex)
        Else
            If cboGrupo.ItemData(cboGrupo.ListIndex) > 0 Then
              xSQL = xSQL & "    AND [Codigo do Grupo] = " & cboGrupo.ItemData(cboGrupo.ListIndex)
            End If
        End If
        xSQL = xSQL & " GROUP BY [Codigo do Produto]"
        xSQL = xSQL & " ORDER BY [Codigo do Produto]"
        Set rsMovimentoVendaConveniencia = Conectar.RsConexao(xSQL)
        If rsMovimentoVendaConveniencia.RecordCount > 0 Then
            Do Until rsMovimentoVendaConveniencia.EOF
                If CalculaEstoqueSubEstoque(True, True, rsMovimentoVendaConveniencia("Codigo do Produto").Value, 1, rsMovimentoVendaConveniencia("Quantidade").Value, xOperacaoSoma) Then
                Else
'aquiaqui                    xAtualizou = False
'                    Conectar.CancelaTransacao
'                    MsgBox "Todo o processamento de cálculo da saída conveniêcia está cancelado!", vbCritical, "Erro de Integridade!"
'                    Call GravaAuditoria(1, Me.name, 26, "Todo o processamento de cálculo da saída conveniêcia está cancelado!")
'                    Exit Do
                End If
'                If Estoque.LocalizarCodigo(g_empresa, rsMovimentoVendaConveniencia("Codigo do Produto").Value) Then
'                    Estoque.Quantidade = Estoque.Quantidade - rsMovimentoVendaConveniencia("Quantidade").Value
'                    If Estoque.Alterar(g_empresa, rsMovimentoVendaConveniencia("Codigo do Produto").Value) Then
'                        If SubEstoque.LocalizarCodigo(g_empresa, rsMovimentoVendaConveniencia("Codigo do Produto").Value, 1) Then
'                            SubEstoque.Quantidade = SubEstoque.Quantidade - rsMovimentoVendaConveniencia("Quantidade").Value
'                            If Not SubEstoque.Alterar(g_empresa, rsMovimentoVendaConveniencia("Codigo do Produto").Value, 1) Then
'                                xAtualizou = False
'                                Conectar.CancelaTransacao
'                                MsgBox "Não foi possível alterar o Sub-Estoque!", vbInformation, "Erro de Integridade!"
'                                Exit Do
'                            End If
'                        End If
'                    Else
'                        xAtualizou = False
'                        Conectar.CancelaTransacao
'                        MsgBox "Não foi possível alterar o estoque!", vbInformation, "Erro de Integridade!"
'                        Exit Do
'                    End If
'                End If
                rsMovimentoVendaConveniencia.MoveNext
            Loop
        End If
        If xAtualizou Then
            Conectar.ConfirmaTransacao
        End If
        rsMovimentoVendaConveniencia.Close
        Set rsMovimentoVendaConveniencia = Nothing
        MsgBox "A empresa " & g_nome_empresa & Chr(10) & "Está com as saidas de conveniencia calculadas para o estoque.", vbInformation, "Operação Concluida!"
    End If
    Exit Sub

trata_erro:
    Conectar.CancelaTransacao
    MsgBox Err.Number & " - " & Err.Description
End Sub
Private Sub ProcessamentoCalculaSaidaInterna()
Dim rsMovimentoLubrificante As adodb.Recordset
Dim xSQL As String
Dim xAtualizou As Boolean
Dim xOperacaoSoma As Boolean

On Error GoTo trata_erro
    
    If (MsgBox("Na empresa " & g_nome_empresa & Chr(10) & "Será calculado as saidas internas de " & msk_data_inicial.Text & " até " & msk_data_final.Text & " para todo estoque." & Chr(10) & Chr(10) & "Deseja realmente fazer este processamento?", vbYesNo + 256, "Calcula Saidas Internas para o Estoque!")) = 6 Then
        Call GravaAuditoria(1, Me.name, 26, "Saídas Interna Empresa:" & g_empresa & "-" & Mid(g_nome_empresa, 1, 20) & " De:" & msk_data_inicial.Text & " a " & msk_data_final.Text)
        xOperacaoSoma = False
        If chkCalculoInverso.Value = 1 Then
            xOperacaoSoma = True
        End If
        
        Set rsMovimentoLubrificante = New adodb.Recordset
        xAtualizou = True
        Conectar.IniciaTransacao
        xSQL = ""
        'xSQL = xSQL & "SELECT [Codigo do Produto2], Quantidade, [Codigo do Tipo do SubEstoque]" ..old 21/07/2015
        xSQL = xSQL & "SELECT [Codigo do Produto2], [Codigo do Tipo do SubEstoque], SUM (Quantidade) AS Quantidade" 'new 21/07/2015
        xSQL = xSQL & "  FROM Movimento_Lubrificante"
        xSQL = xSQL & " WHERE Empresa = " & g_empresa
        xSQL = xSQL & "   AND Data >= " & preparaData(CDate(msk_data_inicial.Text))
        xSQL = xSQL & "   AND Data <= " & preparaData(CDate(msk_data_final.Text))
        If cboProduto.ItemData(cboProduto.ListIndex) > 0 Then
            xSQL = xSQL & "    AND [Codigo do Produto2] = " & cboProduto.ItemData(cboProduto.ListIndex)
        Else
            If cboGrupo.ItemData(cboGrupo.ListIndex) > 0 Then
                xSQL = xSQL & " AND [Codigo do Produto2] IN"
                xSQL = xSQL & "   ("
                xSQL = xSQL & "    SELECT Produto.Codigo"
                xSQL = xSQL & "      FROM Produto"
                xSQL = xSQL & "     WHERE Produto.[Codigo do Grupo] = " & cboGrupo.ItemData(cboGrupo.ListIndex)
                xSQL = xSQL & "   )"
            End If
        End If
        'xSQL = xSQL & " ORDER BY [Codigo do Produto2], Data" ..old 21/07/2015
        xSQL = xSQL & " GROUP BY [Codigo do Produto2], [Codigo do Tipo do SubEstoque]" 'new 21/07/2015
        xSQL = xSQL & " ORDER BY [Codigo do Produto2], [Codigo do Tipo do SubEstoque]" 'new 21/07/2015
        
        Set rsMovimentoLubrificante = Conectar.RsConexao(xSQL)
        If rsMovimentoLubrificante.RecordCount > 0 Then
            Do Until rsMovimentoLubrificante.EOF
                If CalculaEstoqueSubEstoque(True, True, rsMovimentoLubrificante("Codigo do Produto2").Value, 1, rsMovimentoLubrificante("Quantidade").Value, xOperacaoSoma) Then
                    If Me.chkTransferenciaSubEstoque.Value = 1 Then
                        'Diminui no Sub-Estoque de Saída
                        If CalculaEstoqueSubEstoque(False, True, rsMovimentoLubrificante("Codigo do Produto2").Value, rsMovimentoLubrificante("Codigo do Tipo do SubEstoque").Value, rsMovimentoLubrificante("Quantidade").Value, xOperacaoSoma) Then
                        Else
                            xAtualizou = False
                            Conectar.CancelaTransacao
                            MsgBox "Todo o processamento de cálculo da saída interna está cancelado!", vbCritical, "Erro de Integridade!"
                            Call GravaAuditoria(1, Me.name, 26, "Todo o processamento de cálculo da saída interna(SubEstoque) está cancelado!")
                            Exit Do
                        End If
                    End If
                Else
                    xAtualizou = False
                    Conectar.CancelaTransacao
                    MsgBox "Todo o processamento de cálculo da saída interna está cancelado!", vbCritical, "Erro de Integridade!"
                    Call GravaAuditoria(1, Me.name, 26, "Todo o processamento de cálculo da saída interna está cancelado!")
                    Exit Do
                End If
'                If Estoque.AlterarQuantidade(g_empresa, rsMovimentoLubrificante("Codigo do Produto2").Value, rsMovimentoLubrificante("Quantidade").Value, False) Then
'                    If Me.chkTransferenciaSubEstoque.Value = 0 Then
'                        'Diminui no Sub-Estoque do Depósito
'                        If Not SubEstoque.AlterarQuantidade(g_empresa, rsMovimentoLubrificante("Codigo do Produto2").Value, 1, rsMovimentoLubrificante("Quantidade").Value, False) Then
'
'                            'obs este código foi usado no posto Costa
'                            'Porque a empresa = 2
'                            'SubEstoque.Empresa = g_empresa
'                            'SubEstoque.CodigoProduto = rsMovimentoLubrificante("Codigo do Produto2").Value
'                            'SubEstoque.CodigoTipoSubEstoque = 1
'                            'SubEstoque.Quantidade = 0
'                            'SubEstoque.Incluir
'                            'SubEstoque.CodigoTipoSubEstoque = 2
'                            'SubEstoque.Incluir
'
'
'                            xAtualizou = False
'                            Conectar.CancelaTransacao
'                            MsgBox "Não foi possível alterar o Sub-Estoque!", vbInformation, "Erro de Integridade!"
'                            Exit Do
'                        End If
'                    End If
'                    If Me.chkTransferenciaSubEstoque.Value = 1 Then
'                        'Diminui no Sub-Estoque de Saída
'                        If Not SubEstoque.AlterarQuantidade(g_empresa, rsMovimentoLubrificante("Codigo do Produto2").Value, rsMovimentoLubrificante("Codigo do Tipo do SubEstoque").Value, rsMovimentoLubrificante("Quantidade").Value, False) Then
'                            xAtualizou = False
'                            Conectar.CancelaTransacao
'                            MsgBox "Não foi possível alterar o Sub-Estoque!", vbInformation, "Erro de Integridade!"
'                            Exit Do
'                        End If
'                    End If
'                Else
''                    xAtualizou = False
''                    Conectar.CancelaTransacao
''                    MsgBox "Não foi possível alterar o estoque!", vbInformation, "Erro de Integridade!"
''                    Exit Do
'                End If
                rsMovimentoLubrificante.MoveNext
            Loop
        End If
        rsMovimentoLubrificante.Close
        Set rsMovimentoLubrificante = Nothing
        
        
        If xAtualizou Then
            Set rsMovimentoLubrificante = New adodb.Recordset
            xSQL = ""
            xSQL = xSQL & "SELECT [Codigo do Produto2], Quantidade"
            xSQL = xSQL & "  FROM Saida_Transferencia_Produto"
            xSQL = xSQL & " WHERE Empresa = " & g_empresa
            xSQL = xSQL & "   AND [Data da Transferencia] >= " & preparaData(CDate(msk_data_inicial.Text))
            xSQL = xSQL & "   AND [Data da Transferencia] <= " & preparaData(CDate(msk_data_final.Text))
            xSQL = xSQL & " ORDER BY [Codigo do Produto2], [Data da Transferencia]"
            Set rsMovimentoLubrificante = Conectar.RsConexao(xSQL)
            If rsMovimentoLubrificante.RecordCount > 0 Then
                Do Until rsMovimentoLubrificante.EOF
                    If Estoque.LocalizarCodigo(g_empresa, rsMovimentoLubrificante("Codigo do Produto2").Value) Then
                        Estoque.Quantidade = Estoque.Quantidade - rsMovimentoLubrificante("Quantidade").Value
                        If Not Estoque.Alterar(g_empresa, rsMovimentoLubrificante("Codigo do Produto2").Value) Then
                            xAtualizou = False
                            Conectar.CancelaTransacao
                            MsgBox "Não foi possível alterar o estoque!", vbInformation, "Erro de Integridade!"
                            Exit Do
                        End If
                    End If
                    rsMovimentoLubrificante.MoveNext
                Loop
            End If
            If xAtualizou Then
                Conectar.ConfirmaTransacao
            End If
            rsMovimentoLubrificante.Close
            Set rsMovimentoLubrificante = Nothing
        End If
        MsgBox "A empresa " & g_nome_empresa & Chr(10) & "Está com as saidas internas calculadas para o estoque.", vbInformation, "Operação Concluida!"
    End If
    Exit Sub

trata_erro:
    Conectar.CancelaTransacao
    MsgBox Err.Number & " - " & Err.Description
End Sub
Private Sub ProcessamentoSomaEntradaInventario()
Dim rsEntradaProduto As adodb.Recordset
Dim xSQL As String
Dim xAtualizou As Boolean

On Error GoTo trata_erro
    
    If (MsgBox("Na empresa " & g_nome_empresa & Chr(10) & "Será somado as entradas (inventário do estoque inicial do dia) da data " & msk_data_inicial.Text & " no o estoque." & Chr(10) & Chr(10) & "Deseja realmente fazer este processamento?", vbYesNo + 256, "Soma Entradas do Inventário no Estoque!")) = vbYes Then
        xAtualizou = True
        Call GravaAuditoria(1, Me.name, 26, "Soma Ent(Invent)no Est.Empresa:" & g_empresa & "-" & Mid(g_nome_empresa, 1, 20) & " Dt:" & msk_data_inicial.Text)
        Set rsEntradaProduto = New adodb.Recordset
        Conectar.IniciaTransacao
        xSQL = ""
        'xSQL = xSQL & "SELECT [Codigo do Produto], Quantidade, SubEstoque" 'old 21/07/2015
        xSQL = xSQL & "SELECT [Codigo do Produto], SubEstoque, SUM(Quantidade) AS Quantidade" 'new 21/07/2015
        xSQL = xSQL & "  FROM Entrada_Produto"
        xSQL = xSQL & " WHERE Empresa = " & g_empresa
        xSQL = xSQL & "   AND [Data da Entrada] >= " & preparaData(CDate(msk_data_inicial.Text))
        xSQL = xSQL & "   AND [Data da Entrada] <= " & preparaData(CDate(msk_data_final.Text))
        'xSQL = xSQL & "   AND [Numero do Documento] = " & preparaTexto("1")
        xSQL = xSQL & "   AND [Tipo da Entrada] = " & preparaTexto("3")
        If cboProduto.ItemData(cboProduto.ListIndex) > 0 Then 'new 21/07/2015 daqui...
            xSQL = xSQL & "    AND [Codigo do Produto] = " & cboProduto.ItemData(cboProduto.ListIndex)
        Else
            If cboGrupo.ItemData(cboGrupo.ListIndex) > 0 Then
                xSQL = xSQL & " AND [Codigo do Produto] IN"
                xSQL = xSQL & "   ("
                xSQL = xSQL & "    SELECT Produto.Codigo"
                xSQL = xSQL & "      FROM Produto"
                xSQL = xSQL & "     WHERE Produto.[Codigo do Grupo] = " & cboGrupo.ItemData(cboGrupo.ListIndex)
                xSQL = xSQL & "   )"
            End If
        End If '... ate aqui new 21/07/2015
        'xSQL = xSQL & " ORDER BY [Codigo do Produto], [Data da Entrada]" 'old 21/07/2015
        xSQL = xSQL & " GROUP BY [Codigo do Produto], SubEstoque" 'new 21/07/2015
        xSQL = xSQL & " ORDER BY [Codigo do Produto], SubEstoque" 'new 21/07/2015
        Set rsEntradaProduto = Conectar.RsConexao(xSQL)
        If rsEntradaProduto.RecordCount > 0 Then
            Do Until rsEntradaProduto.EOF
                If CalculaEstoqueSubEstoque(True, True, rsEntradaProduto("Codigo do Produto").Value, rsEntradaProduto("SubEstoque").Value, rsEntradaProduto("Quantidade").Value, True) Then
                Else
'                    xAtualizou = False
'                    Conectar.CancelaTransacao
'aquiaqui                    MsgBox "Todo o processamento de cálculo de inventário de entrada está cancelado!", vbCritical, "Erro de Integridade!"
'                    Call GravaAuditoria(1, Me.name, 26, "Todo o processamento de cálculo de inventário de entrada está cancelado!")
'                    Exit Do
                End If
'                If Estoque.LocalizarCodigo(g_empresa, rsEntradaProduto("Codigo do Produto").Value) Then
'                    Estoque.Quantidade = rsEntradaProduto("Quantidade").Value
'                    If Estoque.Alterar(g_empresa, rsEntradaProduto("Codigo do Produto").Value) Then
'                        If SubEstoque.LocalizarCodigo(g_empresa, rsEntradaProduto("Codigo do Produto").Value, rsEntradaProduto("SubEstoque").Value) Then
'                            SubEstoque.Quantidade = rsEntradaProduto("Quantidade").Value
'                            If Not SubEstoque.Alterar(g_empresa, rsEntradaProduto("Codigo do Produto").Value, rsEntradaProduto("SubEstoque").Value) Then
'                                xAtualizou = False
'                                Conectar.CancelaTransacao
'                                MsgBox "Não foi possível alterar o Sub-Estoque!", vbInformation, "Erro de Integridade!"
'                                Exit Do
'                            End If
'                        End If
'                    Else
'                        xAtualizou = False
'                        Conectar.CancelaTransacao
'                        MsgBox "Não foi possível alterar o estoque!", vbInformation, "Erro de Integridade!"
'                        Exit Do
'                    End If
'                End If
                rsEntradaProduto.MoveNext
            Loop
        End If
        If xAtualizou Then
            Conectar.ConfirmaTransacao
        End If
        rsEntradaProduto.Close
        Set rsEntradaProduto = Nothing
        MsgBox "A empresa " & g_nome_empresa & Chr(10) & "Está com as entradas (inventário) movidas para o estoque.", vbInformation, "Operação Concluida!"
    End If
    Exit Sub

trata_erro:
    Conectar.CancelaTransacao
    MsgBox Err.Number & " - " & Err.Description
End Sub
Private Sub ProcessamentoMoveInventarioContabil()
Dim rsEstoque2 As adodb.Recordset
Dim xSQL As String
Dim xAtualizou As Boolean

On Error GoTo trata_erro
    
    If (MsgBox("Na empresa " & g_nome_empresa & Chr(10) & "Será movido o inventário Contábil da data " & msk_data_inicial & " para todo estoque." & Chr(10) & Chr(10) & "Deseja realmente fazer este processamento?", vbYesNo + 256, "Move Entradas para o Estoque!")) = vbYes Then
        Call GravaAuditoria(1, Me.name, 26, "Move Invent.Contábil p/Est.Empresa:" & g_empresa & "-" & Mid(g_nome_empresa, 1, 20) & " Dt:" & msk_data_inicial.Text)
        Set rsEstoque2 = New adodb.Recordset
        xAtualizou = True
        Conectar.IniciaTransacao
        xSQL = ""
        xSQL = xSQL & "SELECT [Codigo do Produto2], Quantidade"
        xSQL = xSQL & "  FROM Estoque2"
        xSQL = xSQL & " WHERE Empresa = " & g_empresa
        xSQL = xSQL & "   AND Data = " & preparaData(CDate(msk_data_inicial.Text))
        If cboGrupo.ItemData(cboGrupo.ListIndex) > 0 Then
            xSQL = xSQL & "    AND [Grupo do Produto] = " & cboGrupo.ItemData(cboGrupo.ListIndex)
        End If
        If cboProduto.ItemData(cboProduto.ListIndex) > 0 Then
            xSQL = xSQL & "    AND [Codigo do Produto2] = " & cboProduto.ItemData(cboProduto.ListIndex)
        End If
        xSQL = xSQL & " ORDER BY [Codigo do Produto2]"
        Set rsEstoque2 = Conectar.RsConexao(xSQL)
        If rsEstoque2.RecordCount > 0 Then
            Do Until rsEstoque2.EOF
                If CalculaEstoqueSubEstoque(True, True, rsEstoque2("Codigo do Produto2").Value, 1, rsEstoque2("Quantidade").Value, True) Then
                Else
'                    xAtualizou = False
'                    Conectar.CancelaTransacao
'                    MsgBox "Todo o processamento de cálculo de inventário contábil está cancelado!", vbCritical, "Erro de Integridade!"
'                    Call GravaAuditoria(1, Me.name, 26, "Todo o processamento de cálculo de inventário contábil está cancelado!")
'                    Exit Do
                End If
'                If Estoque.LocalizarCodigo(g_empresa, rsEstoque2("Codigo do Produto2").Value) Then
'                    Estoque.Quantidade = rsEstoque2("Quantidade").Value
'                    If Estoque.Alterar(g_empresa, rsEstoque2("Codigo do Produto2").Value) Then
'                        If SubEstoque.LocalizarCodigo(g_empresa, rsEstoque2("Codigo do Produto2").Value, 1) Then
'                            SubEstoque.Quantidade = rsEstoque2("Quantidade").Value
'                            If Not SubEstoque.Alterar(g_empresa, rsEstoque2("Codigo do Produto2").Value, 1) Then
'                                xAtualizou = False
'                                Conectar.CancelaTransacao
'                                MsgBox "Não foi possível alterar o Sub-Estoque!", vbInformation, "Erro de Integridade!"
'                                Exit Do
'                            End If
'                        End If
'                    Else
'                        xAtualizou = False
'                        Conectar.CancelaTransacao
'                        MsgBox "Não foi possível alterar o estoque!", vbInformation, "Erro de Integridade!"
'                        Exit Do
'                    End If
'                End If
                rsEstoque2.MoveNext
            Loop
        End If
        If xAtualizou Then
            Conectar.ConfirmaTransacao
        End If
        rsEstoque2.Close
        Set rsEstoque2 = Nothing
        MsgBox "A empresa " & g_nome_empresa & Chr(10) & "Está com o inventário Contábil movido para o estoque atual.", vbInformation, "Operação Concluida!"
    End If
    Exit Sub

trata_erro:
    Conectar.CancelaTransacao
    MsgBox Err.Number & " - " & Err.Description
End Sub
Private Sub ProcessamentoTransfSubEstoque()
Dim rsTransfInternaEstoque As adodb.Recordset
Dim xSQL As String
Dim xAtualizou As Boolean
Dim xOperacaoSoma As Boolean

On Error GoTo trata_erro
    
    If (MsgBox("Na empresa " & g_nome_empresa & Chr(10) & "Será calculado as transferências internas de " & msk_data_inicial & " até " & msk_data_final & " para todo Sub-Estoque." & Chr(10) & Chr(10) & "Deseja realmente fazer este processamento?", vbYesNo + 256, "Calcula Transf. Interna!")) = vbYes Then
        Call GravaAuditoria(1, Me.name, 26, "Calcula transf.interna Empresa:" & g_empresa & "-" & Mid(g_nome_empresa, 1, 20) & " De:" & msk_data_inicial.Text & " a " & msk_data_final.Text)
        
        Set rsTransfInternaEstoque = New adodb.Recordset
        xAtualizou = True
        Conectar.IniciaTransacao
        xSQL = ""
        xSQL = xSQL & "SELECT [Codigo do Produto], [Codigo do SubEstoque de Saida], "
        xSQL = xSQL & "       [Codigo do SubEstoque de Entrada], Quantidade"
        xSQL = xSQL & "  FROM TransferenciaInternaEstoque"
        xSQL = xSQL & " WHERE Empresa = " & g_empresa
        xSQL = xSQL & "   AND Data >= " & preparaData(CDate(msk_data_inicial.Text))
        xSQL = xSQL & "   AND Data <= " & preparaData(CDate(msk_data_final.Text))
        xSQL = xSQL & "   AND Transferido = " & preparaBooleano(True)
        xSQL = xSQL & " ORDER BY Data, [Codigo do Produto]"
        Set rsTransfInternaEstoque = Conectar.RsConexao(xSQL)
        If rsTransfInternaEstoque.RecordCount > 0 Then
            Do Until rsTransfInternaEstoque.EOF
                
                
                'SubEstoque p/ Saída
                xOperacaoSoma = False
                If chkCalculoInverso.Value = 1 Then
                    xOperacaoSoma = True
                End If
                If CalculaEstoqueSubEstoque(False, True, rsTransfInternaEstoque("Codigo do Produto").Value, rsTransfInternaEstoque("Codigo do SubEstoque de Saida").Value, rsTransfInternaEstoque("Quantidade").Value, xOperacaoSoma) Then
                Else
                    xAtualizou = False
                    Conectar.CancelaTransacao
                    MsgBox "Todo o processamento de cálculo da Saída do SubEstoque está cancelado!", vbCritical, "Erro de Integridade!"
                    Call GravaAuditoria(1, Me.name, 26, "Todo o processamento de cálculo da Saída do SubEstoque está cancelado!")
                    Exit Do
                End If
'                If SubEstoque.LocalizarCodigo(g_empresa, rsTransfInternaEstoque("Codigo do Produto").Value, rsTransfInternaEstoque("Codigo do SubEstoque de Saida").Value) Then
'                    SubEstoque.Quantidade = SubEstoque.Quantidade - rsTransfInternaEstoque("Quantidade").Value
'                    If Not SubEstoque.Alterar(g_empresa, rsTransfInternaEstoque("Codigo do Produto").Value, rsTransfInternaEstoque("Codigo do SubEstoque de Saida").Value) Then
'                        xAtualizou = False
'                        Conectar.CancelaTransacao
'                        MsgBox "Não foi possível alterar o Sub-Estoque Saída!", vbInformation, "Erro de Integridade!"
'                        Exit Do
'                    End If
'                End If
                
                
                'SubEstoque p/ Entrada
                xOperacaoSoma = True
                If chkCalculoInverso.Value = 1 Then
                    xOperacaoSoma = False
                End If
                If CalculaEstoqueSubEstoque(False, True, rsTransfInternaEstoque("Codigo do Produto").Value, rsTransfInternaEstoque("Codigo do SubEstoque de Entrada").Value, rsTransfInternaEstoque("Quantidade").Value, xOperacaoSoma) Then
                Else
                    xAtualizou = False
                    Conectar.CancelaTransacao
                    MsgBox "Todo o processamento de cálculo da Entrada do SubEstoque está cancelado!", vbCritical, "Erro de Integridade!"
                    Call GravaAuditoria(1, Me.name, 26, "Todo o processamento de cálculo da Entrada do SubEstoque está cancelado!")
                    Exit Do
                End If
'                If SubEstoque.LocalizarCodigo(g_empresa, rsTransfInternaEstoque("Codigo do Produto").Value, rsTransfInternaEstoque("Codigo do SubEstoque de Entrada").Value) Then
'                    SubEstoque.Quantidade = SubEstoque.Quantidade + rsTransfInternaEstoque("Quantidade").Value
'                    If Not SubEstoque.Alterar(g_empresa, rsTransfInternaEstoque("Codigo do Produto").Value, rsTransfInternaEstoque("Codigo do SubEstoque de Entrada").Value) Then
'                        xAtualizou = False
'                        Conectar.CancelaTransacao
'                        MsgBox "Não foi possível alterar o Sub-Estoque Entrada!", vbInformation, "Erro de Integridade!"
'                        Exit Do
'                    End If
'                End If
                rsTransfInternaEstoque.MoveNext
            Loop
        End If
        If xAtualizou Then
            Conectar.ConfirmaTransacao
        End If
        rsTransfInternaEstoque.Close
        Set rsTransfInternaEstoque = Nothing
        
        MsgBox "A empresa " & g_nome_empresa & Chr(10) & "Está com as transferências internas calculadas para todo Sub-Estoque.", vbInformation, "Operação Concluida!"
    End If
    Exit Sub

trata_erro:
    Conectar.CancelaTransacao
    MsgBox Err.Number & " - " & Err.Description
End Sub
Private Sub ProcessamentoZeraEstoque()
Dim xSQL As String
On Error GoTo trata_erro
    
    If (MsgBox("Na empresa " & g_nome_empresa & Chr(10) & "Será feito o processamento para zerar todo seu estoque." & Chr(10) & Chr(10) & "Deseja realmente fazer este processamento?", vbYesNo + 256, "Zeramento de Estoque!")) = 6 Then
        Call GravaAuditoria(1, Me.name, 26, "Zera Estoque Atual - Empresa:" & g_empresa & "-" & g_nome_empresa)
        Conectar.IniciaTransacao
        xSQL = ""
        xSQL = xSQL & "UPDATE Estoque"
        xSQL = xSQL & "   SET Quantidade = 0"
        xSQL = xSQL & " WHERE Empresa = " & g_empresa
        If cboGrupo.ItemData(cboGrupo.ListIndex) > 0 Then
            xSQL = xSQL & "    AND [Grupo do Produto] = " & cboGrupo.ItemData(cboGrupo.ListIndex)
        End If
        If cboProduto.ItemData(cboProduto.ListIndex) > 0 Then
            xSQL = xSQL & "    AND [Codigo do Produto2] = " & cboProduto.ItemData(cboProduto.ListIndex)
        End If
        Conectar.ExecutaSql (xSQL)
        
        
        xSQL = ""
        If cboGrupo.ItemData(cboGrupo.ListIndex) > 0 And cboProduto.ItemData(cboProduto.ListIndex) = 0 Then
            xSQL = xSQL & "UPDATE SubEstoque"
            xSQL = xSQL & "   SET Quantidade = 0"
            xSQL = xSQL & " WHERE [Codigo do Produto] IN"
            xSQL = xSQL & "   ("
            xSQL = xSQL & "    SELECT SubEstoque.[Codigo do Produto]"
            xSQL = xSQL & "      FROM SubEstoque, Produto"
            xSQL = xSQL & "     WHERE Empresa = " & g_empresa
            xSQL = xSQL & "       AND SubEstoque.[Codigo do Produto] = Produto.Codigo"
            xSQL = xSQL & "       AND Produto.[Codigo do Grupo] = " & cboGrupo.ItemData(cboGrupo.ListIndex)
            xSQL = xSQL & "     GROUP BY SubEstoque.[codigo do produto]"
            xSQL = xSQL & "   )"
        Else
            xSQL = xSQL & "UPDATE SubEstoque"
            xSQL = xSQL & "   SET Quantidade = 0"
            xSQL = xSQL & " WHERE Empresa = " & g_empresa
            If cboProduto.ItemData(cboProduto.ListIndex) > 0 Then
                xSQL = xSQL & "    AND [Codigo do Produto] = " & cboProduto.ItemData(cboProduto.ListIndex)
            End If
        End If
        Conectar.ExecutaSql (xSQL)
        Conectar.ConfirmaTransacao
        MsgBox "A empresa " & g_nome_empresa & Chr(10) & "Está com todo seu estoque zerado.", vbInformation, "Operação Concluida!"
    End If
    Exit Sub

trata_erro:
    Conectar.CancelaTransacao
    MsgBox Err.Number & " - " & Err.Description
End Sub
Private Sub cboGrupo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cboProduto.SetFocus
    End If
End Sub
Private Sub cboProduto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub chk_calcula_entrada_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        chk_calcula_saida.SetFocus
    End If
End Sub
Private Sub chk_calcula_saida_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data_inicial.SetFocus
    End If
End Sub
Private Sub chkSomaEntradaInventario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        chkMoveInventarioContabil.SetFocus
    End If
End Sub
Private Sub chk_zera_estoque_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        chkSomaEntradaInventario.SetFocus
    End If
End Sub
Private Sub chkMoveInventarioContabil_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        chk_calcula_entrada.SetFocus
    End If
End Sub
Private Sub cmd_ok_Click()
    On Error GoTo FileError
    If ValidaCampos Then
        Processamento
        cmd_sair.SetFocus
    End If
    Exit Sub
FileError:
    MsgBox "Erro no processamento!", vbInformation, "Erro Interno"
    Exit Sub
End Sub
Private Sub PreencheCboGrupo()
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "   SELECT Codigo, Nome"
    lSQL = lSQL & "     FROM Grupo"
    lSQL = lSQL & " ORDER BY Nome, Codigo"
    'Abre RecordSet
    Set rs = New adodb.Recordset
    Set rs = Conectar.RsConexao(lSQL)
    
    cboGrupo.Clear
    cboGrupo.AddItem "Todos os Grupos"
    cboGrupo.ItemData(cboGrupo.NewIndex) = 0
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        Do Until rs.EOF
            cboGrupo.AddItem rs("Nome").Value
            cboGrupo.ItemData(cboGrupo.NewIndex) = rs("Codigo").Value
            rs.MoveNext
        Loop
    End If
End Sub
Private Sub PreencheCboProduto()
    cboProduto.Clear
    
    cboProduto.AddItem "Todos os Produtos"
    cboProduto.ItemData(cboProduto.NewIndex) = 0
    lSQL = ""
    lSQL = lSQL & "SELECT Codigo, Nome"
    lSQL = lSQL & "  FROM Produto"
    lSQL = lSQL & " ORDER BY Nome, Codigo"
    Set rs = Conectar.RsConexao(lSQL)
    If rs.RecordCount > 0 Then
        Do Until rs.EOF
                cboProduto.AddItem rs!Nome
                cboProduto.ItemData(cboProduto.NewIndex) = rs!Codigo
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
End Sub
Function ValidaCampos() As Boolean
    ValidaCampos = False
    If Not IsDate(msk_data_inicial.Text) Then
        MsgBox "Informe a data inicial.", vbInformation, "Atenção!"
        msk_data_inicial.SetFocus
    ElseIf Not IsDate(msk_data_final.Text) Then
        MsgBox "Informe a data final.", vbInformation, "Atenção!"
        msk_data_final.SetFocus
    ElseIf Not IsDate(msk_data_final.Text) >= IsDate(msk_data_inicial.Text) Then
        MsgBox "A data final deve ser igual ou maior que " & msk_data_inicial.Text & " .", vbInformation, "Atenção!"
        msk_data_final.SetFocus
    ElseIf chk_zera_estoque.Value = 0 And chkSomaEntradaInventario.Value = 0 And chkMoveInventarioContabil.Value = 0 And chk_calcula_entrada.Value = 0 And chk_calcula_saida.Value = 0 And Me.chkTransferenciaSubEstoque.Value = 0 Then
        MsgBox "Deve ser selecionada pelo menos uma das opções acima.", vbInformation, "Atenção!"
        chk_zera_estoque.SetFocus
    ElseIf cboGrupo.ListIndex = -1 Then
        MsgBox "Selecione o grupo.", vbInformation, "Atenção!"
        cboGrupo.SetFocus
    ElseIf cboProduto.ListIndex = -1 Then
        MsgBox "Selecione o produto.", vbInformation, "Atenção!"
        cboProduto.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    Screen.MousePointer = 1
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        cmd_sair_Click
    ElseIf KeyCode = vbKeyF11 Then
        KeyCode = 0
        cmd_ok_Click
    End If
End Sub
Private Sub Form_Load()
    CentraForm Me
    
    PreencheCboGrupo
    PreencheCboProduto
    msk_data_inicial.Text = Format(g_data_def, "dd/mm/yyyy")
    msk_data_final.Text = Format(g_data_def, "dd/mm/yyyy")
    chk_zera_estoque.Value = 0
    chkSomaEntradaInventario.Value = 1
    chk_calcula_entrada.Value = 1
    chk_calcula_saida.Value = 1
    chkVendaInterna.Value = 1
    cboGrupo.ListIndex = 0
    cboProduto.ListIndex = 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_final_GotFocus()
    msk_data_final.SelStart = 0
    msk_data_final.SelLength = 5
End Sub
Private Sub msk_data_final_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub msk_data_inicial_GotFocus()
    msk_data_inicial.SelStart = 0
    msk_data_inicial.SelLength = 5
End Sub
Private Sub msk_data_inicial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data_final.SetFocus
    End If
End Sub
