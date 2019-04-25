VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form processamento_custo_combustivel 
   Caption         =   "Processamento de Custo de Combustível"
   ClientHeight    =   3690
   ClientLeft      =   165
   ClientTop       =   285
   ClientWidth     =   6495
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   Picture         =   "processamento_custo_combustivel.frx":0000
   ScaleHeight     =   3690
   ScaleWidth      =   6495
   Begin VB.Frame frmDados 
      Height          =   2595
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   6255
      Begin VB.CheckBox chkMovimentoPista 
         Caption         =   "Movimento de Pista"
         Height          =   315
         Left            =   300
         TabIndex        =   10
         Top             =   240
         Width           =   2775
      End
      Begin VB.OptionButton optCustoReal 
         Caption         =   "Custo Real"
         Height          =   255
         Left            =   300
         TabIndex        =   1
         Top             =   660
         Value           =   -1  'True
         Width           =   2595
      End
      Begin VB.OptionButton optCustoMedio 
         Caption         =   "Custo Médio"
         Height          =   255
         Left            =   300
         TabIndex        =   3
         Top             =   1380
         Width           =   2595
      End
      Begin VB.OptionButton optCustoRealMedio 
         Caption         =   "Custo Real Médio"
         Height          =   255
         Left            =   300
         TabIndex        =   2
         Top             =   1020
         Width           =   2595
      End
      Begin MSMask.MaskEdBox msk_data_inicial 
         Height          =   300
         Left            =   180
         TabIndex        =   5
         Top             =   2100
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
         Left            =   2820
         TabIndex        =   7
         Top             =   2100
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
         TabIndex        =   4
         Top             =   1890
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "D&ata final"
         Height          =   195
         Index           =   8
         Left            =   2820
         TabIndex        =   6
         Top             =   1890
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   1680
      Picture         =   "processamento_custo_combustivel.frx":0446
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Confirma o processamento."
      Top             =   2760
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4020
      Picture         =   "processamento_custo_combustivel.frx":1A50
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   2760
      Width           =   795
   End
End
Attribute VB_Name = "processamento_custo_combustivel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lSQL As String
Private rsTabela As New adodb.Recordset

Private EntradaCombustivel As New cEntradaCombustivel
Private LivroLMC As New cLivroLMC
Private MedicaoCombustivel As New cMedicaoCombustivel
Private MovAfericao As New cMovimentoAfericao
Private MovimentoBomba As New cMovimentoBomba
Private PrecoCustoCombustivel As New cPrecoCustoCombustivel
Private Sub AtivaBotoes(ByVal pAtiva As Boolean)
    cmd_ok.Enabled = pAtiva
    cmd_sair.Enabled = pAtiva
    If pAtiva = False Then
        frmAguarde.Show
        Call frmAguarde.MostraMensagens("Processando!", Me.Caption, Me.Top, Me.Left, Me.Width, Me.Height)
        DoEvents
    Else
        Call frmAguarde.Finaliza
    End If
End Sub
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set EntradaCombustivel = Nothing
    Set LivroLMC = Nothing
    Set MovAfericao = Nothing
    Set MedicaoCombustivel = Nothing
    Set MovimentoBomba = Nothing
    Set PrecoCustoCombustivel = Nothing
End Sub
Private Sub Processamento()
    If optCustoReal.Value = True Then
        ProcessamentoCustoReal
    End If
    If optCustoRealMedio.Value = True Then
        ProcessamentoCustoRealMedio
    End If
    If optCustoMedio.Value = True Then
        ProcessamentoCustoMedio
    End If
End Sub
Private Sub ProcessamentoCustoMedio()
    Dim xData As Date
    Dim xTipoCombustivel As String
    Dim xQtdCusto As Currency
    Dim xValorCusto As Currency
    Dim i As Integer
    If (MsgBox("Na empresa " & g_nome_empresa & Chr(10) & "Será calculado o custo médio de combustível entre " & msk_data_inicial.Text & " a " & msk_data_final.Text & "." & Chr(10) & Chr(10) & "Deseja realmente fazer este processamento?", vbYesNo + 256, "Calcula Custo Médio de Combustível!")) = vbYes Then
        Call GravaAuditoria(1, Me.name, 26, "Calcula Custo de Combustível: Custo Médio")
        Call GravaAuditoria(2, Me.name, 26, "Empresa:" & g_empresa & "-" & Mid(g_nome_empresa, 1, 20) & " De:" & msk_data_inicial.Text & " a " & msk_data_final.Text)
        'Prepara SQL
        lSQL = ""
        lSQL = lSQL & "SELECT Nome, Codigo"
        lSQL = lSQL & "  FROM Combustivel"
        lSQL = lSQL & " ORDER BY Nome"
        'Abre RecordSet
        Set rsTabela = Conectar.RsConexao(lSQL)
        'Verifica Registros
        If rsTabela.RecordCount > 0 Then
            Do Until rsTabela.EOF
                xQtdCusto = 0
                xValorCusto = 0
                xTipoCombustivel = rsTabela("Codigo").Value
                For xData = CDate(msk_data_inicial.Text) To CDate(msk_data_final.Text)
                    g_string = EntradaCombustivel.DadosEntradaData(g_empresa, xData, xTipoCombustivel)
                    If Val(RetiraGString(1)) > 0 Then
                        xQtdCusto = xQtdCusto + fValidaValor(RetiraGString(1))
                        xValorCusto = xValorCusto + fValidaValor(RetiraGString(2)) * fValidaValor(RetiraGString(1))
                    End If
                Next
                If xQtdCusto > 0 Then
                    xValorCusto = Format(xValorCusto / xQtdCusto, "00000000.0000")
                    For xData = CDate(msk_data_inicial.Text) To CDate(msk_data_final.Text)
                        If Not MovimentoBomba.AlteraPrecoCusto(g_empresa, xData, xTipoCombustivel, xValorCusto) Then
                            MsgBox "Não possível alterar o preço de custo!", vbInformation, "Erro de integridade."
                        End If
                        If MovAfericao.TotalQtdPeriodoCombustivel(g_empresa, xData, xData, 1, 9, xTipoCombustivel, "") > 0 Then
                            If Not MovAfericao.AlteraPrecoCusto(g_empresa, xData, xTipoCombustivel, PrecoCustoCombustivel.ValorVenda) Then
                                MsgBox "Não possível alterar o preço de custo de Aferição!", vbInformation, "Erro de integridade."
                            End If
                        End If
                    Next
                End If
                rsTabela.MoveNext
            Loop
        End If
        If rsTabela.State = 1 Then
            rsTabela.Close
        End If
        MsgBox "A empresa " & g_nome_empresa & Chr(10) & "Está com o custo médio calculado.", vbInformation, "Processamento Concluído!"
    End If
End Sub
Private Sub ProcessamentoCustoReal()
    Dim xData As Date
    Dim xPreco As Currency
    Dim xValor(0 To 10) As Currency
    Dim xQuantidade(0 To 10)  As Currency
    Dim xQtdVendaDia As Currency
    Dim xQtdEntradaDia As Currency
    Dim xQtdDeQuantidade As Integer
    Dim xTipoCombustivel As String
    Dim xString As String
    Dim xQtdCusto As Currency
    Dim xValorCusto As Currency
    Dim i As Integer
    Dim xSequencia As String
    
    On Error GoTo Error_ProcessamentoCustoReal
    
    If (MsgBox("Na empresa " & g_nome_empresa & Chr(10) & "Será calculado o custo real de combustível entre " & msk_data_inicial.Text & " a " & msk_data_final.Text & "." & Chr(10) & Chr(10) & "Deseja realmente fazer este processamento?", vbYesNo + 256, "Calcula Custo Real de Combustível!")) = vbYes Then
        Call GravaAuditoria(1, Me.name, 26, "Calcula Custo de Combustível: Custo Real")
        Call GravaAuditoria(2, Me.name, 26, "Empresa:" & g_empresa & "-" & Mid(g_nome_empresa, 1, 20) & " De:" & msk_data_inicial.Text & " a " & msk_data_final.Text)
        lSQL = ""
        lSQL = lSQL & "DELETE"
        lSQL = lSQL & "  FROM Preco_Custo_Combustivel"
        lSQL = lSQL & " WHERE Empresa = " & g_empresa
        lSQL = lSQL & "   AND Data >= " & preparaData(CDate(msk_data_inicial.Text))
        lSQL = lSQL & "   AND Data <= " & preparaData(CDate(msk_data_final.Text))
        xSequencia = "1"
        Conectar.ExecutaSql (lSQL)
        xSequencia = "2"
        'Prepara SQL
        lSQL = ""
        lSQL = lSQL & "SELECT Nome, Codigo"
        lSQL = lSQL & "  FROM Combustivel"
        lSQL = lSQL & " WHERE Empresa = " & g_empresa
        lSQL = lSQL & " ORDER BY Nome"
        'Abre RecordSet
        Set rsTabela = Conectar.RsConexao(lSQL)
        xSequencia = "3"
        'Verifica Registros
        If rsTabela.RecordCount > 0 Then
            Do Until rsTabela.EOF
                xTipoCombustivel = rsTabela("Codigo").Value
                xSequencia = "4"
                If MedicaoCombustivel.LocalizarPrimeiroTanqueComb(g_empresa, CDate(msk_data_inicial.Text), xTipoCombustivel) Then
                    xSequencia = "5"
                    xQuantidade(0) = MedicaoCombustivel.TotalMedidaCombustivel(g_empresa, CDate(msk_data_inicial.Text), xTipoCombustivel, 0)
                    xSequencia = "6"
                    xQtdDeQuantidade = 1
                    xValor(0) = 0
                    If EntradaCombustivel.LocalizarUltimoCombustivel(g_empresa, CDate(msk_data_inicial.Text), xTipoCombustivel) Then
                        xSequencia = "7"
                        xValor(0) = EntradaCombustivel.ValorLitro
                    End If
                    xSequencia = "8"
                    PrecoCustoCombustivel.Empresa = g_empresa
                    PrecoCustoCombustivel.TipoCombustivel = xTipoCombustivel
                    PrecoCustoCombustivel.Data = CDate(msk_data_inicial.Text) - 1
                    PrecoCustoCombustivel.Ordem = 1
                    PrecoCustoCombustivel.QuantidadeInicial = xQuantidade(0)
                    PrecoCustoCombustivel.ValorInicial = xValor(0)
                    PrecoCustoCombustivel.QuantidadeEntrada = 0
                    PrecoCustoCombustivel.ValorEntrada = 0
                    PrecoCustoCombustivel.QuantidadeVenda = 0
                    PrecoCustoCombustivel.ValorVenda = 0
                    PrecoCustoCombustivel.QuantidadePrecoAnterior = 0
                    xSequencia = "9"
                    If Not PrecoCustoCombustivel.Incluir("CustoReal") Then
                        xSequencia = "10"
                        If Not PrecoCustoCombustivel.Alterar(g_empresa, xTipoCombustivel, CDate(msk_data_inicial.Text) - 1, 1) Then
                            xSequencia = "11"
                            MsgBox "Erro ao alterar preco de custo de combustivel", vbInformation
                        End If
                    End If
                    xSequencia = "12"
                    For xData = CDate(msk_data_inicial.Text) To CDate(msk_data_final.Text)
                        xQtdCusto = 0
                        xValorCusto = 0
                        xSequencia = "13"
                        xQtdVendaDia = MovimentoBomba.QuantidadeVendaData(g_empresa, xData, xData, xTipoCombustivel, 0)
                        xSequencia = "14"
                        If PrecoCustoCombustivel.LocalizarCodigo(g_empresa, xTipoCombustivel, (xData - 1), 1) Then
                            xQuantidade(0) = PrecoCustoCombustivel.QuantidadeInicial + PrecoCustoCombustivel.QuantidadeEntrada - PrecoCustoCombustivel.QuantidadeVenda
                            xSequencia = "15"
                            If PrecoCustoCombustivel.QuantidadeEntrada > 0 Then
                                xValor(0) = ((PrecoCustoCombustivel.ValorInicial * PrecoCustoCombustivel.QuantidadeInicial) + (PrecoCustoCombustivel.ValorEntrada * PrecoCustoCombustivel.QuantidadeEntrada)) / (PrecoCustoCombustivel.QuantidadeInicial + PrecoCustoCombustivel.QuantidadeEntrada)
                                xSequencia = "16"
                            Else
                                xValor(0) = PrecoCustoCombustivel.ValorInicial
                                xSequencia = "17"
                            End If
                            PrecoCustoCombustivel.Data = xData
                            PrecoCustoCombustivel.Ordem = 1
                            PrecoCustoCombustivel.QuantidadeInicial = xQuantidade(0)
                            PrecoCustoCombustivel.ValorInicial = xValor(0)
                            PrecoCustoCombustivel.QuantidadeEntrada = 0
                            PrecoCustoCombustivel.ValorEntrada = 0
                            PrecoCustoCombustivel.QuantidadeVenda = xQtdVendaDia
                            PrecoCustoCombustivel.ValorVenda = 0
                            PrecoCustoCombustivel.QuantidadePrecoAnterior = 0
                            xSequencia = "18"
                            If (PrecoCustoCombustivel.QuantidadeInicial - PrecoCustoCombustivel.QuantidadeVenda) > 0 Then
                                xSequencia = "19"
                                PrecoCustoCombustivel.QuantidadePrecoAnterior = (PrecoCustoCombustivel.QuantidadeInicial - PrecoCustoCombustivel.QuantidadeVenda)
                            End If
                            xSequencia = "20"
                            g_string = EntradaCombustivel.DadosEntradaData(g_empresa, xData, xTipoCombustivel)
                            xSequencia = "21"
                            If Val(RetiraGString(1)) > 0 Then
                                PrecoCustoCombustivel.QuantidadeEntrada = fValidaValor(RetiraGString(1))
                                PrecoCustoCombustivel.ValorEntrada = fValidaValor(RetiraGString(2))
                                xSequencia = "22"
                                'PrecoCustoCombustivel.ValorInicial = fValidaValor(RetiraGString(2))
                                'xValor(0) = Format(((xValor(0) * xQuantidade(0)) + (fValidaValor(RetiraGString(1)) * fValidaValor(RetiraGString(2)))) / (xQuantidade(0) + fValidaValor(RetiraGString(1))), "000000.0000")
                                'xQuantidade(0) = xQuantidade(0) + Val(RetiraGString(1))
                            End If
                            g_string = ""
                            If Not PrecoCustoCombustivel.Incluir("CustoReal") Then
                                MsgBox "Erro ao incluir preco de custo de combustivel", vbInformation
                            End If
                            xSequencia = "23"
                            
                            'If xQtdVendaDia > xQuantidade(0) Then
                            '    xQtdCusto = xQuantidade(0)
                            '    xValorCusto = xValor(0)
                            '    xQtdVendaDia = xQtdVendaDia - xQuantidade(0)
                            '    xQuantidade(0) = 0
                            '    g_string = EntradaCombustivel.DadosEntradaData(g_empresa, xData, xTipoCombustivel)
                            '    If Val(RetiraGString(1)) > 0 Then
                            '        xValorCusto = Format(((xQtdCusto * xValorCusto) + (xQtdVendaDia * fValidaValor(RetiraGString(2)))) / (xQtdCusto + xQtdVendaDia), "000000.0000")
                            '        xQtdVendaDia = xQtdVendaDia + xQtdVendaDia
                            '        xValor(0) = fValidaValor(RetiraGString(2))
                            '        xQuantidade(0) = fValidaValor(RetiraGString(1)) - xQtdVendaDia
                            '    End If
                            '    g_string = ""
                           '
                           '
                           '     PrecoCustoCombustivel.Data = xData
                           '     PrecoCustoCombustivel.Ordem = i
                           '     PrecoCustoCombustivel.QuantidadeInicial = xQuantidade(0)
                           '     PrecoCustoCombustivel.ValorInicial = xValor(0)
                           '     If Not PrecoCustoCombustivel.Incluir Then
                           '         MsgBox "Erro ao incluir preco de custo de combustivel", vbInformation
                           '     End If
                           '
                           '
                            'Else
                            '    xQtdCusto = xQtdVendaDia
                            '    xValorCusto = xValor(0)
                            '    xQuantidade(0) = xQuantidade(0) - xQtdVendaDia
                            '    g_string = EntradaCombustivel.DadosEntradaData(g_empresa, xData, xTipoCombustivel)
                            '    If Val(RetiraGString(1)) > 0 Then
                            '        xValor(0) = Format(((xValor(0) * xQuantidade(0)) + (fValidaValor(RetiraGString(1)) * fValidaValor(RetiraGString(2)))) / (xQuantidade(0) + fValidaValor(RetiraGString(1))), "000000.0000")
                            '        xQuantidade(0) = xQuantidade(0) + Val(RetiraGString(1))
                            '    End If
                            '    g_string = ""
                           '
                           '     PrecoCustoCombustivel.Data = xData
                           '     PrecoCustoCombustivel.Ordem = i
                           '     PrecoCustoCombustivel.QuantidadeInicial = xQuantidade(0)
                           '     PrecoCustoCombustivel.ValorInicial = xValor(0)
                           '     If Not PrecoCustoCombustivel.Incluir Then
                           '         MsgBox "Erro ao incluir preco de custo de combustivel", vbInformation
                           '     End If
                            
                            
                            'End If
                            If Not MovimentoBomba.AlteraPrecoCusto(g_empresa, xData, xTipoCombustivel, PrecoCustoCombustivel.ValorVenda) Then
                                MsgBox "Não foi possível alterar o preço de custo no movimento de bomba!" & vbCrLf & "Data: " & xData & vbCrLf & "Tipo Combustível: " & xTipoCombustivel, vbInformation, "Erro de integridade."
                            End If
                            If MovAfericao.TotalQtdPeriodoCombustivel(g_empresa, xData, xData, 1, 9, xTipoCombustivel, "") > 0 Then
                                xSequencia = "24"
                                If Not MovAfericao.AlteraPrecoCusto(g_empresa, xData, xTipoCombustivel, PrecoCustoCombustivel.ValorVenda) Then
                                    MsgBox "Não possível alterar o preço de custo de Aferição!", vbInformation, "Erro de integridade."
                                End If
                            End If
                        Else
                            MsgBox "erro"
                        End If
                    Next
                Else
                    MsgBox "Falta estoque inicial do combustível " & xTipoCombustivel & Chr(10) & "Na data " & msk_data_inicial.Text & ".", vbInformation, "Combustível sem Medida!"
                    'Exit Sub
                End If
                rsTabela.MoveNext
            Loop
        End If
        If rsTabela.State = 1 Then
            rsTabela.Close
        End If
        MsgBox "A empresa " & g_nome_empresa & Chr(10) & "Está com o custo real calculado.", vbInformation, "Processamento Concluído!"
    End If
    Exit Sub

Error_ProcessamentoCustoReal:
    MsgBox "Erro: " & Error & Chr(10) & "Erro N.: " & Err & Chr(10) & "Sequancia: " & xSequencia, vbCritical, "Erro Não Identificado!"
    If rsTabela.State = 1 Then
        rsTabela.Close
    End If
End Sub
Private Sub ProcessamentoCustoRealMedio()
    Dim xData As Date
    Dim xPreco As Currency
    Dim xValor(0 To 10) As Currency
    Dim xQuantidade(0 To 10)  As Currency
    Dim xQtdVendaDia As Currency
    Dim xQtdEntradaDia As Currency
    Dim xQtdDeQuantidade As Integer
    Dim xTipoCombustivel As String
    Dim xString As String
    Dim xQtdCusto As Currency
    Dim xValorCusto As Currency
    Dim i As Integer
    
    If (MsgBox("Na empresa " & g_nome_empresa & Chr(10) & "Será calculado o custo real de combustível entre " & msk_data_inicial.Text & " a " & msk_data_final.Text & "." & Chr(10) & Chr(10) & "Deseja realmente fazer este processamento?", vbYesNo + 256, "Calcula Custo Real de Combustível!")) = vbYes Then
        Call GravaAuditoria(1, Me.name, 26, "Calcula Custo de Combustível: Custo Real Médio")
        Call GravaAuditoria(2, Me.name, 26, "Empresa:" & g_empresa & "-" & Mid(g_nome_empresa, 1, 20) & " De:" & msk_data_inicial.Text & " a " & msk_data_final.Text)
        lSQL = ""
        lSQL = lSQL & "DELETE"
        lSQL = lSQL & "  FROM Preco_Custo_Combustivel"
        lSQL = lSQL & " WHERE Empresa = " & g_empresa
        lSQL = lSQL & "   AND Data >= " & preparaData(CDate(msk_data_inicial.Text))
        lSQL = lSQL & "   AND Data <= " & preparaData(CDate(msk_data_final.Text))
        Conectar.ExecutaSql (lSQL)
        'Prepara SQL
        lSQL = ""
        lSQL = lSQL & "SELECT Nome, Codigo"
        lSQL = lSQL & "  FROM Combustivel"
        lSQL = lSQL & " ORDER BY Nome"
        'Abre RecordSet
        Set rsTabela = Conectar.RsConexao(lSQL)
        'Verifica Registros
        If rsTabela.RecordCount > 0 Then
            Do Until rsTabela.EOF
                xTipoCombustivel = rsTabela("Codigo").Value
                If MedicaoCombustivel.LocalizarPrimeiroTanqueComb(g_empresa, CDate(msk_data_inicial.Text), xTipoCombustivel) Then
                    xQuantidade(0) = MedicaoCombustivel.TotalMedidaCombustivel(g_empresa, CDate(msk_data_inicial.Text), xTipoCombustivel, 0)
                    xQtdDeQuantidade = 1
                    xValor(0) = 0
                    If EntradaCombustivel.LocalizarUltimoCombustivel(g_empresa, CDate(msk_data_inicial.Text), xTipoCombustivel) Then
                        xValor(0) = EntradaCombustivel.ValorLitro
                    End If
                    PrecoCustoCombustivel.Empresa = g_empresa
                    PrecoCustoCombustivel.TipoCombustivel = xTipoCombustivel
                    PrecoCustoCombustivel.Data = CDate(msk_data_inicial.Text) - 1
                    PrecoCustoCombustivel.Ordem = 1
                    PrecoCustoCombustivel.QuantidadeInicial = xQuantidade(0)
                    PrecoCustoCombustivel.ValorInicial = xValor(0)
                    PrecoCustoCombustivel.QuantidadeEntrada = 0
                    PrecoCustoCombustivel.ValorEntrada = 0
                    PrecoCustoCombustivel.QuantidadeVenda = 0
                    PrecoCustoCombustivel.ValorVenda = 0
                    PrecoCustoCombustivel.QuantidadePrecoAnterior = 0
                    If Not PrecoCustoCombustivel.Incluir("CustoRealMedio") Then
                        If Not PrecoCustoCombustivel.Alterar(g_empresa, xTipoCombustivel, CDate(msk_data_inicial.Text) - 1, 1) Then
                            MsgBox "Erro ao alterar preco de custo de combustivel", vbInformation
                        End If
                    End If
                    For xData = CDate(msk_data_inicial.Text) To CDate(msk_data_final.Text)
                        xQtdCusto = 0
                        xValorCusto = 0
                        xQtdVendaDia = MovimentoBomba.QuantidadeVendaData(g_empresa, xData, xData, xTipoCombustivel, 0)
                        If PrecoCustoCombustivel.LocalizarCodigo(g_empresa, xTipoCombustivel, (xData - 1), 1) Then
                            xQuantidade(0) = PrecoCustoCombustivel.QuantidadeInicial + PrecoCustoCombustivel.QuantidadeEntrada - PrecoCustoCombustivel.QuantidadeVenda
                            If PrecoCustoCombustivel.QuantidadeEntrada > 0 Then
                                xValor(0) = ((PrecoCustoCombustivel.ValorInicial * PrecoCustoCombustivel.QuantidadeInicial) + (PrecoCustoCombustivel.ValorEntrada * PrecoCustoCombustivel.QuantidadeEntrada)) / (PrecoCustoCombustivel.QuantidadeInicial + PrecoCustoCombustivel.QuantidadeEntrada)
                            Else
                                xValor(0) = PrecoCustoCombustivel.ValorInicial
                            End If
                            
                            
                            g_string = EntradaCombustivel.DadosEntradaData(g_empresa, xData, xTipoCombustivel)
                            If Val(RetiraGString(1)) > 0 Then
                                PrecoCustoCombustivel.QuantidadeEntrada = fValidaValor(RetiraGString(1))
                                PrecoCustoCombustivel.ValorEntrada = fValidaValor(RetiraGString(2))
                                'PrecoCustoCombustivel.ValorInicial = fValidaValor(RetiraGString(2))
                                xValor(0) = Format(((xValor(0) * xQuantidade(0)) + (fValidaValor(RetiraGString(1)) * fValidaValor(RetiraGString(2)))) / (xQuantidade(0) + fValidaValor(RetiraGString(1))), "000000.0000")
                                xQuantidade(0) = xQuantidade(0) + Val(RetiraGString(1))
                            End If
                            g_string = ""
                            
                            
                            
                            PrecoCustoCombustivel.Data = xData
                            PrecoCustoCombustivel.Ordem = 1
                            PrecoCustoCombustivel.QuantidadeInicial = xQuantidade(0)
                            PrecoCustoCombustivel.ValorInicial = xValor(0)
                            PrecoCustoCombustivel.QuantidadeEntrada = 0
                            PrecoCustoCombustivel.ValorEntrada = 0
                            PrecoCustoCombustivel.QuantidadeVenda = xQtdVendaDia
                            PrecoCustoCombustivel.ValorVenda = 0
                            PrecoCustoCombustivel.QuantidadePrecoAnterior = 0
                            If (PrecoCustoCombustivel.QuantidadeInicial - PrecoCustoCombustivel.QuantidadeVenda) > 0 Then
                                PrecoCustoCombustivel.QuantidadePrecoAnterior = (PrecoCustoCombustivel.QuantidadeInicial - PrecoCustoCombustivel.QuantidadeVenda)
                            End If
                            If Not PrecoCustoCombustivel.Incluir("CustoRealMedio") Then
                                MsgBox "Erro ao incluir preco de custo de combustivel", vbInformation
                            End If
                            If Not MovimentoBomba.AlteraPrecoCusto(g_empresa, xData, xTipoCombustivel, PrecoCustoCombustivel.ValorVenda) Then
                                MsgBox "Não possível alterar o preço de custo!", vbInformation, "Erro de integridade."
                            End If
                            If MovAfericao.TotalQtdPeriodoCombustivel(g_empresa, xData, xData, 1, 9, xTipoCombustivel, "") > 0 Then
                                If Not MovAfericao.AlteraPrecoCusto(g_empresa, xData, xTipoCombustivel, PrecoCustoCombustivel.ValorVenda) Then
                                    MsgBox "Não possível alterar o preço de custo de Aferição!", vbInformation, "Erro de integridade."
                                End If
                            End If
                        Else
                            MsgBox "erro"
                        End If
                    Next
                Else
                    MsgBox "Falta estoque inicial da data " & msk_data_inicial.Text & ".", vbInformation, "Erro de Verificação!"
                    Exit Sub
                End If
                rsTabela.MoveNext
            Loop
        End If
        If rsTabela.State = 1 Then
            rsTabela.Close
        End If
        MsgBox "A empresa " & g_nome_empresa & Chr(10) & "Está com o custo real calculado.", vbInformation, "Processamento Concluído!"
    End If
End Sub
Private Sub cmd_ok_Click()
    On Error GoTo FileError
    If ValidaCampos Then
        AtivaBotoes (False)
        If chkMovimentoPista.Visible = True Then
            If chkMovimentoPista.Value = 1 Then
                MovimentoBomba.NomeTabela = "Movimento_Bomba_Cupom"
            Else
                MovimentoBomba.NomeTabela = "Movimento_Bomba"
            End If
        End If
        Processamento
        AtivaBotoes (True)
        cmd_sair.SetFocus
    End If
    Exit Sub
FileError:
    MsgBox "Erro " & Error & Chr(10) & "Erro Numero:" & Err, vbCritical, "Erro Nao Identificado!"
    'ErroArquivo tbl_estoque.Name, "Estoqueo"
    Exit Sub
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
    ElseIf Not ValidaLiberacaoLMC(CDate(msk_data_inicial.Text)) Then
        msk_data_inicial.SetFocus
    ElseIf Not ValidaLiberacaoLMC(CDate(msk_data_final.Text)) Then
        msk_data_final.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Function ValidaLiberacaoLMC(ByVal pData As Date) As Boolean
    ValidaLiberacaoLMC = False
    
    If g_nome_usuario = "L.M.C." Then
        If LivroLMC.LocalizarCombustivelConcluido(g_empresa, "**", pData) = "NAO" Then
            ValidaLiberacaoLMC = True
        ElseIf LivroLMC.LocalizarCombustivelConcluido(g_empresa, "**", pData) = "SIM" Then
            MsgBox "O LMC está concluído nesta data.", vbCritical, "LMC concluído!"
        ElseIf LivroLMC.LocalizarCombustivelConcluido(g_empresa, "**", pData) = "**" Then
            MsgBox "O LMC não está com páginas cadastradas.", vbCritical, "Página não cadastrada!"
        End If
    Else
        ValidaLiberacaoLMC = True
    End If
End Function
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub Form_Activate()
    If cmd_ok.Enabled Then
        cmd_ok.SetFocus
    End If
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
    Call GravaAuditoria(1, Me.name, 1, "")
    Screen.MousePointer = 1
    CentraForm Me
    
    If g_nome_usuario = "L.M.C." Then
        chkMovimentoPista.Visible = False
        EntradaCombustivel.NomeTabela = "Entrada_Combustivel_LMC"
        MedicaoCombustivel.NomeTabela = "MedicaoCombustivelLMC"
        MovAfericao.NomeTabela = "Movimento_Afericao_LMC"
        MovimentoBomba.NomeTabela = "Movimento_Bomba_LMC"
    Else
        EntradaCombustivel.NomeTabela = "Entrada_Combustivel"
        MedicaoCombustivel.NomeTabela = "MedicaoCombustivel"
        MovAfericao.NomeTabela = "Movimento_Afericao"
        If ReadINI("CUPOM FISCAL", "ECF Instalada", gArquivoIni) = "SIM" Then
            chkMovimentoPista.Value = 1
            MovimentoBomba.NomeTabela = "Movimento_Bomba_Cupom"
        Else
            chkMovimentoPista.Value = 0
            MovimentoBomba.NomeTabela = "Movimento_Bomba"
        End If
    End If
    
    msk_data_inicial.Text = fDataPrimeiroDiaMesAnterior(Date)
    msk_data_final.Text = fDataUltimoDiaMesAnterior(Date)
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
Private Sub optCustoMedio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data_inicial.SetFocus
    End If
End Sub
Private Sub optCustoReal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data_inicial.SetFocus
    End If
End Sub
Private Sub optCustoRealMedio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data_inicial.SetFocus
    End If
End Sub
