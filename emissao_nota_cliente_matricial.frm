VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form emissao_nota_cliente_matricial 
   Caption         =   "Emissão das Notas de Abastecimento por Cliente (Matricial)"
   ClientHeight    =   4455
   ClientLeft      =   3990
   ClientTop       =   2010
   ClientWidth     =   7530
   Icon            =   "emissao_nota_cliente_matricial.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   Picture         =   "emissao_nota_cliente_matricial.frx":030A
   ScaleHeight     =   4455
   ScaleWidth      =   7530
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1440
      Picture         =   "emissao_nota_cliente_matricial.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Visualiza notas de abastecimento por emissão."
      Top             =   3540
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3300
      Picture         =   "emissao_nota_cliente_matricial.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Imprime notas de abastecimento por emissão."
      Top             =   3540
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   5160
      Picture         =   "emissao_nota_cliente_matricial.frx":3074
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   3540
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   3315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7275
      Begin VB.CheckBox chkImprimirNumeroNota 
         Caption         =   "Imprimir Número da Nota de Abastecimento"
         Height          =   255
         Left            =   1680
         TabIndex        =   26
         Top             =   2940
         Width           =   5355
      End
      Begin VB.CheckBox chkNotaConferida 
         Caption         =   "&Notas Conferidas"
         Height          =   255
         Left            =   4260
         TabIndex        =   4
         Top             =   240
         Width           =   2235
      End
      Begin VB.TextBox txtDataFinal 
         Height          =   285
         Left            =   5160
         MaxLength       =   10
         TabIndex        =   12
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtDataInicial 
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   9
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtDataEmissao 
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox chkSomenteNotaBaixada 
         Caption         =   "Somente Baixadas"
         Height          =   255
         Left            =   4920
         TabIndex        =   22
         Top             =   2640
         Width           =   2235
      End
      Begin VB.CheckBox chkNotaBaixada 
         Caption         =   "Imprimir Notas já Baixadas"
         Height          =   255
         Left            =   1680
         TabIndex        =   21
         Top             =   2640
         Width           =   2835
      End
      Begin VB.CheckBox chkValorLiquido 
         Caption         =   "Imprimir Val&or Líquido"
         Height          =   255
         Left            =   4260
         TabIndex        =   16
         Top             =   1440
         Width           =   2835
      End
      Begin VB.ComboBox cboGrupo 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1800
         Width           =   5475
      End
      Begin VB.ComboBox cboProduto 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   2220
         Width           =   5475
      End
      Begin VB.CheckBox chkUnificaEmpresa 
         Height          =   255
         Left            =   1680
         TabIndex        =   15
         Top             =   1440
         Width           =   435
      End
      Begin VB.TextBox txt_cliente 
         Height          =   285
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   6
         Top             =   660
         Width           =   795
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2820
         Picture         =   "emissao_nota_cliente_matricial.frx":4706
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2820
         Picture         =   "emissao_nota_cliente_matricial.frx":59E0
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   6300
         Picture         =   "emissao_nota_cliente_matricial.frx":6CBA
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   1080
         Width           =   495
      End
      Begin MSAdodcLib.Adodc adodcCliente 
         Height          =   330
         Left            =   4020
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
      Begin MSDataListLib.DataCombo dtcboCliente 
         Bindings        =   "emissao_nota_cliente_matricial.frx":7F94
         Height          =   315
         Left            =   2580
         TabIndex        =   7
         Top             =   660
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Razao Social"
         BoundColumn     =   "Codigo"
         Text            =   "dtcboCliente"
      End
      Begin VB.Label Label3 
         Caption         =   "&Grupo"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   1800
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "P&roduto"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   19
         Top             =   2220
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "&Unifica empresas"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "C&liente"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   5
         Top             =   660
         Width           =   1515
      End
      Begin VB.Label Label5 
         Caption         =   "Data de &emissão"
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1515
      End
      Begin VB.Label Label2 
         Caption         =   "D&ata final"
         Height          =   315
         Left            =   4200
         TabIndex        =   11
         Top             =   1080
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "&Data inicial"
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   1515
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   3900
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "emissao_nota_cliente_matricial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Início de variáveis padrão para relatório
Dim lLinha As Integer
Dim lPagina As Integer
Dim lLocal As Integer
Dim lNomeArquivo As String
'Fim de variáveis padrão para relatório
Dim lTotalQtd As Currency
Dim lTotal As Currency
Dim lTotalDesconto As Currency
Dim lTotalQtdDesconto As Currency
Dim lSQL As String
Dim rstMovimentoNota As adodb.Recordset
Dim rsTabela As adodb.Recordset

Private Cliente As New cCliente
Private MovimentoCupomFiscal As New cMovimentoCupomFiscal
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set Cliente = Nothing
    Set MovimentoCupomFiscal = Nothing
End Sub
Private Sub ZeraVariaveis()
    lLinha = 0
    lPagina = 0
    lTotalQtd = 0
    lTotal = 0
    lTotalQtdDesconto = 0
    lTotalDesconto = 0
End Sub
Private Sub PreencheCboGrupo()
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "   SELECT Codigo, Nome"
    lSQL = lSQL & "     FROM Grupo"
    lSQL = lSQL & " ORDER BY Nome, Codigo"
    'Abre RecordSet
    Set rsTabela = New adodb.Recordset
    Set rsTabela = Conectar.RsConexao(lSQL)
    
    cboGrupo.Clear
    cboGrupo.AddItem "Todos os Grupos"
    cboGrupo.ItemData(cboGrupo.NewIndex) = 0
    If rsTabela.RecordCount > 0 Then
        rsTabela.MoveFirst
        Do Until rsTabela.EOF
            cboGrupo.AddItem rsTabela("Nome").Value
            cboGrupo.ItemData(cboGrupo.NewIndex) = rsTabela("Codigo").Value
            rsTabela.MoveNext
        Loop
    End If
    rsTabela.Close
    Set rsTabela = Nothing
End Sub
Private Sub PreencheCboProduto()
    cboProduto.Clear
    
    cboProduto.AddItem "Todos os Produtos"
    cboProduto.ItemData(cboProduto.NewIndex) = 0
    lSQL = ""
    lSQL = lSQL & "SELECT Codigo, Nome"
    lSQL = lSQL & "  FROM Produto"
    lSQL = lSQL & " WHERE [Exclusivo Posto] = " & preparaBooleano(True)
    lSQL = lSQL & " ORDER BY Nome, Codigo"
    Set rsTabela = New adodb.Recordset
    Set rsTabela = Conectar.RsConexao(lSQL)
    If rsTabela.RecordCount > 0 Then
        Do Until rsTabela.EOF
                cboProduto.AddItem rsTabela!Nome
                cboProduto.ItemData(cboProduto.NewIndex) = rsTabela!Codigo
            rsTabela.MoveNext
        Loop
    End If
    rsTabela.Close
    Set rsTabela = Nothing
End Sub
Private Sub Relatorio()
    Dim xExisteNota As Boolean
    
    ZeraVariaveis
    xExisteNota = False
    If chkSomenteNotaBaixada.Value = 0 Then
        lSQL = ""
        lSQL = lSQL & " SELECT [Numero da Nota], [Data do Abastecimento], [Valor Total], [Valor Unitario],"
        lSQL = lSQL & "        Quantidade, [Valor Desconto Unitario], [Codigo do Produto2], Empresa,"
        lSQL = lSQL & "        Periodo, Produto.Nome as NomeProduto"
        If chkImprimirNumeroNota.Value = 1 Then
            lSQL = lSQL & ", Origem, [Numero do Cupom], Ordem"
        End If
        lSQL = lSQL & "   FROM Movimento_Nota_Abastecimento, Produto"
        lSQL = lSQL & "  WHERE [Codigo do Cliente] = " & CLng(txt_cliente.Text)
        If chkNotaConferida.Value = 1 Then
            lSQL = lSQL & "    AND [Data da Conferencia] >= " & preparaData(CDate(txtDataInicial.Text))
            lSQL = lSQL & "    AND [Data da Conferencia] <= " & preparaData(CDate(txtDataFinal.Text))
        Else
            lSQL = lSQL & "    AND [Data do Abastecimento] >= " & preparaData(CDate(txtDataInicial.Text))
            lSQL = lSQL & "    AND [Data do Abastecimento] <= " & preparaData(CDate(txtDataFinal.Text))
        End If
        If chkUnificaEmpresa.Value = 0 Then
            lSQL = lSQL & "    AND Empresa = " & g_empresa
        End If
        lSQL = lSQL & "    AND [Codigo do Produto2] = Produto.Codigo"
        If cboGrupo.ItemData(cboGrupo.ListIndex) > 0 Then
            lSQL = lSQL & "    AND Produto.[Codigo do Grupo] = " & cboGrupo.ItemData(cboGrupo.ListIndex)
        End If
        If cboProduto.ItemData(cboProduto.ListIndex) > 0 Then
            lSQL = lSQL & "    AND [Codigo do Produto2] = " & cboProduto.ItemData(cboProduto.ListIndex)
        End If
        lSQL = lSQL & "  ORDER BY [Data do Abastecimento], Periodo, [Numero da Nota]"
        Set rstMovimentoNota = Conectar.RsConexao(lSQL)
        If rstMovimentoNota.RecordCount > 0 Then
            xExisteNota = True
            ImpDados
        End If
        rstMovimentoNota.Close
        Set rstMovimentoNota = Nothing
    End If
    If Me.chkNotaBaixada.Value = 1 Then
        lSQL = ""
        lSQL = lSQL & " SELECT [Numero da Nota], [Data do Abastecimento], [Valor Total], [Valor Unitario],"
        lSQL = lSQL & "        Quantidade, [Valor Desconto Unitario], [Codigo do Produto2], Empresa,"
        lSQL = lSQL & "        Periodo, Produto.Nome as NomeProduto"
        If chkImprimirNumeroNota.Value = 1 Then
            lSQL = lSQL & ", Origem, [Numero do Cupom], Ordem"
        End If
        lSQL = lSQL & "   FROM Baixa_Nota_Abastecimento, Produto"
        lSQL = lSQL & "  WHERE [Codigo do Cliente] = " & CLng(txt_cliente.Text)
        If chkNotaConferida.Value = 1 Then
            lSQL = lSQL & "    AND [Data da Conferencia] >= " & preparaData(CDate(txtDataInicial.Text))
            lSQL = lSQL & "    AND [Data da Conferencia] <= " & preparaData(CDate(txtDataFinal.Text))
        Else
            lSQL = lSQL & "    AND [Data do Abastecimento] >= " & preparaData(CDate(txtDataInicial.Text))
            lSQL = lSQL & "    AND [Data do Abastecimento] <= " & preparaData(CDate(txtDataFinal.Text))
        End If
        If chkUnificaEmpresa.Value = 0 Then
            lSQL = lSQL & "    AND Empresa = " & g_empresa
        End If
        lSQL = lSQL & "    AND [Codigo do Produto2] = Produto.Codigo"
        If cboGrupo.ItemData(cboGrupo.ListIndex) > 0 Then
            lSQL = lSQL & "    AND Produto.[Codigo do Grupo] = " & cboGrupo.ItemData(cboGrupo.ListIndex)
        End If
        If cboProduto.ItemData(cboProduto.ListIndex) > 0 Then
            lSQL = lSQL & "    AND [Codigo do Produto2] = " & cboProduto.ItemData(cboProduto.ListIndex)
        End If
        lSQL = lSQL & "  ORDER BY [Data do Abastecimento], Periodo, [Numero da Nota]"
        Set rstMovimentoNota = Conectar.RsConexao(lSQL)
        If rstMovimentoNota.RecordCount > 0 Then
            xExisteNota = True
            ImpDadosBaixa
        Else
            If lTotal > 0 Then
                ImpTotal
                BioImprime "@@Printer.EndDoc"
                BioFechaImprime
                g_string = lLocal & lNomeArquivo & "|@|Relatório de Nota de Cliente|@|"
                frm_preview.Show 1
            End If
        End If
        rstMovimentoNota.Close
        Set rstMovimentoNota = Nothing
    End If
    If chkSomenteNotaBaixada.Value = 1 And chkNotaBaixada.Value = 0 Then
        lSQL = ""
        lSQL = lSQL & " SELECT [Numero da Nota], [Data do Abastecimento], [Valor Total], [Valor Unitario],"
        lSQL = lSQL & "        Quantidade, [Valor Desconto Unitario], [Codigo do Produto2], Empresa,"
        lSQL = lSQL & "        Periodo, Produto.Nome as NomeProduto"
        lSQL = lSQL & "   FROM Baixa_Nota_Abastecimento, Produto"
        lSQL = lSQL & "  WHERE [Codigo do Cliente] = " & CLng(txt_cliente.Text)
        If chkNotaConferida.Value = 1 Then
            lSQL = lSQL & "    AND [Data da Conferencia] >= " & preparaData(CDate(txtDataInicial.Text))
            lSQL = lSQL & "    AND [Data da Conferencia] <= " & preparaData(CDate(txtDataFinal.Text))
        Else
            lSQL = lSQL & "    AND [Data do Abastecimento] >= " & preparaData(CDate(txtDataInicial.Text))
            lSQL = lSQL & "    AND [Data do Abastecimento] <= " & preparaData(CDate(txtDataFinal.Text))
        End If
        If chkUnificaEmpresa.Value = 0 Then
            lSQL = lSQL & "    AND Empresa = " & g_empresa
        End If
        lSQL = lSQL & "    AND [Codigo do Produto2] = Produto.Codigo"
        If cboGrupo.ItemData(cboGrupo.ListIndex) > 0 Then
            lSQL = lSQL & "    AND Produto.[Codigo do Grupo] = " & cboGrupo.ItemData(cboGrupo.ListIndex)
        End If
        If cboProduto.ItemData(cboProduto.ListIndex) > 0 Then
            lSQL = lSQL & "    AND [Codigo do Produto2] = " & cboProduto.ItemData(cboProduto.ListIndex)
        End If
        lSQL = lSQL & "  ORDER BY [Data do Abastecimento], Periodo, [Numero da Nota]"
        Set rstMovimentoNota = Conectar.RsConexao(lSQL)
        If rstMovimentoNota.RecordCount > 0 Then
            xExisteNota = True
            ImpDadosBaixa
        Else
            If lTotal > 0 Then
                ImpTotal
                BioImprime "@@Printer.EndDoc"
                BioFechaImprime
                g_string = lLocal & lNomeArquivo & "|@|Relatório de Nota de Cliente|@|"
                frm_preview.Show 1
            End If
        End If
        rstMovimentoNota.Close
        Set rstMovimentoNota = Nothing
    End If
    If xExisteNota = False Then
        MsgBox "Cliente não tem notas de abastecimento no período informado!", vbInformation, "Relatório não será impresso!"
    End If
    Call GravaAuditoria(1, Me.name, 7, "Cli:" & CLng(txt_cliente.Text) & " Ref:" & txtDataInicial.Text & " a " & txtDataFinal.Text & " Vlr:" & Format(lTotal, "###,###,##0.00"))
    cmd_sair.SetFocus
End Sub
Private Sub ImpDados()
    Dim xLinha As String
    'loop movimento de notas de abastecimento
    With rstMovimentoNota
        Do Until .EOF
            If lPagina = 0 Then
                ImpCab
                'ImpCliente
            End If
            If lLinha >= 60 Then
                xLinha = "+------------+--------+------+------------------------------------------+----------------+----------------+----------------+------------+"
                BioImprime "@Printer.Print " & xLinha
                BioImprime "@@Printer.NewPage"
                ImpCab
            End If
            'If ![Codigo do Cliente] <> l_cliente Or ![Codigo do Conveniado] <> l_conveniado Or ![Numero da Nota] <> l_numero_nota Then
            '    ImpCliente
            'End If
            ImpDet (False)
            If chkValorLiquido.Value = 0 Then
                If ![Valor Desconto Unitario] <> 0 Then
                    lTotalQtdDesconto = lTotalQtdDesconto + !Quantidade
                    lTotalDesconto = lTotalDesconto + Format((!Quantidade * ![Valor Desconto Unitario]), "0000000000.00")
                End If
            End If
            .MoveNext
        Loop
    End With
    If chkNotaBaixada.Value = 0 Then
        If lTotal > 0 Then
            ImpTotal
            BioImprime "@@Printer.EndDoc"
            BioFechaImprime
            g_string = lLocal & lNomeArquivo & "|@|Relatório de Nota de Cliente|@|"
            frm_preview.Show 1
        End If
    End If
End Sub
Private Sub ImpDadosBaixa()
    Dim xLinha As String
    'loop movimento de notas de abastecimento
    With rstMovimentoNota
        Do Until .EOF
            If lPagina = 0 Then
                ImpCab
                'ImpCliente
            End If
            If lLinha >= 60 Then
                xLinha = "+------------+--------+------+------------------------------------------+----------------+----------------+----------------+------------+"
                BioImprime "@Printer.Print " & xLinha
                BioImprime "@@Printer.NewPage"
                ImpCab
            End If
            'If ![Codigo do Cliente] <> l_cliente Or ![Codigo do Conveniado] <> l_conveniado Or ![Numero da Nota] <> l_numero_nota Then
            '    ImpCliente
            'End If
            ImpDet (True)
            If chkValorLiquido.Value = 0 Then
                If ![Valor Desconto Unitario] <> 0 Then
                    lTotalQtdDesconto = lTotalQtdDesconto + !Quantidade
                    lTotalDesconto = lTotalDesconto + Format((!Quantidade * ![Valor Desconto Unitario]), "0000000000.00")
                End If
            End If
            .MoveNext
        Loop
    End With
    If lTotal > 0 Then
        ImpTotal
        BioImprime "@@Printer.EndDoc"
        BioFechaImprime
        g_string = lLocal & lNomeArquivo & "|@|Relatório de Nota de Cliente|@|"
        frm_preview.Show 1
    End If
End Sub
Private Sub ImpDet(ByVal pBaixada As Boolean)
    Dim xLinha As String
    Dim i As Integer
    Dim xValor As Currency
    
    xLinha = Space(137)
    xLinha = "         1         2         3         4         5         6         7         8         9        10        11        12        13     13"
    xLinha = "12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567"
    xLinha = "|12/99/9999| 1212.123 | 1234 | 1234567890123456789012345678901234567890 |   quantidade   |   quantidade   |   quantidade   | empresa    |"
    xLinha = "|          |          |      |                                          |                |                |                |            |"
    With rstMovimentoNota
        If ![Valor Desconto Unitario] <> 0 Then
            'xValor = Format((![Valor Unitario] - ![Valor Desconto Unitario]) * !Quantidade, "0000000000.00")
            xValor = ![Valor Total]
        Else
            xValor = ![Valor Total]
        End If
        If chkValorLiquido.Value = 1 Then
            If ![Valor Desconto Unitario] < 0 Then
                xValor = ![Valor Total] + Format(-![Valor Desconto Unitario] * !Quantidade, "0000000000.00")
            ElseIf ![Valor Desconto Unitario] > 0 Then
                xValor = ![Valor Total] - Format(![Valor Desconto Unitario] * !Quantidade, "0000000000.00")
            End If
        End If
        lTotalQtd = lTotalQtd + !Quantidade
        lTotal = lTotal + xValor
        
        Mid(xLinha, 2, 10) = Format(![Data do Abastecimento], "dd/mm/yyyy")
        i = Len(Format(![Numero da Nota], "#####,##0"))
        Mid(xLinha, 14 + 9 - i, i) = Format(![Numero da Nota], "#####,##0")
        If chkImprimirNumeroNota.Value = 1 Then
            If !Origem = "CF" Then
                If ![Numero do Cupom] > 0 Then
                    If MovimentoCupomFiscal.LocalizarCodigo(g_empresa, 0, ![Numero do Cupom], ![Data do Abastecimento], !Ordem) Then
                        Mid(xLinha, 13, 10) = "          "
                        If Len(MovimentoCupomFiscal.NumeroCheque) > 0 Then
                            i = Len(Format(CLng(MovimentoCupomFiscal.NumeroCheque), "#####,##0"))
                            Mid(xLinha, 14 + 9 - i, i) = Format(CLng(MovimentoCupomFiscal.NumeroCheque), "#####,##0")
                        Else
                            Mid(xLinha, 14 + 9 - i, i) = Format(0, "#####,##0")
                        End If
                    End If
                End If
            End If
        End If
        
        i = Len(Format(![Codigo do Produto2], "#000"))
        Mid(xLinha, 25 + 4 - i, i) = Format(![Codigo do Produto2], "#000")
        Mid(xLinha, 32, 40) = !NomeProduto
        'i = Len(Format(![Valor Unitario], "###,###,##0.00"))
        'Mid(xLinha, 75 + 14 - i, i) = Format(![Valor Unitario], "###,###,##0.00")
        i = Len(Format(!Quantidade, "###,###,##0.00"))
        Mid(xLinha, 92 + 14 - i, i) = Format(!Quantidade, "###,###,##0.00")
        i = Len(Format(xValor, "###,###,##0.00"))
        Mid(xLinha, 109 + 14 - i, i) = Format(xValor, "###,###,##0.00")
        Mid(xLinha, 126, 1) = !Empresa
        Mid(xLinha, 130, 3) = "P-" & !Periodo
        If pBaixada Then
            Mid(xLinha, 134, 2) = "BX"
        End If
    End With
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 1
End Sub
Private Sub ImpTotal()
    Dim xLinha As String
    Dim i As Integer
    
    xLinha = "+----------+----------+------+------------------------------------------+----------------+----------------+----------------+------------+"
    BioImprime "@Printer.Print " & xLinha
    If lTotalDesconto > 0 Then
        xLinha = "|                                                                      Total Bruto       |                |                |            |"
        i = Len(Format(lTotalQtd, "###,###,##0.00"))
        Mid(xLinha, 92 + 14 - i, i) = Format(lTotalQtd, "###,###,##0.00")
        i = Len(Format(lTotal, "###,###,##0.00"))
        Mid(xLinha, 109 + 14 - i, i) = Format(lTotal, "###,###,##0.00")
        BioImprime "@Printer.Print " & xLinha
        xLinha = "|                                                                      (-) Desconto      |                |                |            |"
        i = Len(Format(lTotalQtdDesconto, "###,###,##0.00"))
        Mid(xLinha, 92 + 14 - i, i) = Format(lTotalQtdDesconto, "###,###,##0.00")
        i = Len(Format(lTotalDesconto, "###,###,##0.00"))
        Mid(xLinha, 109 + 14 - i, i) = Format(lTotalDesconto, "###,###,##0.00")
        BioImprime "@Printer.Print " & xLinha
        lTotal = lTotal - lTotalDesconto
    ElseIf lTotalDesconto < 0 Then
        xLinha = "|                                                                      Total Bruto       |                |                |            |"
        i = Len(Format(lTotalQtd, "###,###,##0.00"))
        Mid(xLinha, 92 + 14 - i, i) = Format(lTotalQtd, "###,###,##0.00")
        i = Len(Format(lTotal, "###,###,##0.00"))
        Mid(xLinha, 109 + 14 - i, i) = Format(lTotal, "###,###,##0.00")
        BioImprime "@Printer.Print " & xLinha
        xLinha = "|                                                                      (+) Acrescimo     |                |                |            |"
        lTotal = lTotal - lTotalDesconto
        lTotalDesconto = lTotalDesconto * -1
        i = Len(Format(lTotalQtdDesconto, "###,###,##0.00"))
        Mid(xLinha, 92 + 14 - i, i) = Format(lTotalQtdDesconto, "###,###,##0.00")
        i = Len(Format(lTotalDesconto, "###,###,##0.00"))
        Mid(xLinha, 109 + 14 - i, i) = Format(lTotalDesconto, "###,###,##0.00")
        BioImprime "@Printer.Print " & xLinha
    End If
    xLinha = "|                                                                      Total para acerto |                |                |            |"
    i = Len(Format(lTotalQtd, "###,###,##0.00"))
    Mid(xLinha, 92 + 14 - i, i) = Format(lTotalQtd, "###,###,##0.00")
    i = Len(Format(lTotal, "###,###,##0.00"))
    Mid(xLinha, 109 + 14 - i, i) = Format(lTotal, "###,###,##0.00")
    BioImprime "@Printer.Print " & xLinha
    xLinha = "+----------------------------------------------------------------------------------------+----------------+----------------+------------+"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|                                                                                                                                       |"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|        Declaro que recebi em ____/____/____ todos os originais relacionados neste documento, e autorizo a emissão de nota fiscal      |"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|     bem como sua cobrança bancária.                                                                                                   |"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|                                                                                                                                       |"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|                                 Aceite __________________________________________________                                             |"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|                                                                                                                                       |"
    i = Len(Trim(Cliente.RazaoSocial))
    Mid(xLinha, 47 + (40 - i) / 2, i) = Trim(Cliente.RazaoSocial)
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|                                                                                                                                       |"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "+---------------------------------------------------------------------------------------------------------------------------------------+"
    Mid(xLinha, 5, 22) = " Cerrado Informática. "
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & " "
End Sub
Private Sub ImpCab()
    Dim xLinha As String
    Dim i As Integer
    Dim x_string_40 As String * 40
    If lPagina = 0 Then
        lNomeArquivo = BioCriaImprime
        'seleciona medidas para centímetros
        BioImprime "@@Printer.ScaleMode = 7"
        BioImprime "@@Printer.PaperSize = 1"
        BioImprime "@@Printer.FontName = Draft 10cpi"
        BioImprime "@@Printer.FontName = Draft 10cpi"
        'teste para imprimir letra correta
        BioImprime "@@Printer.FontBold = False"
        BioImprime "@@ImprimeTexto " & Chr(34) & "  " & Chr(34) & ", 1, 2, 2, 1"
    End If
    lPagina = lPagina + 1
    lLinha = 0
    BioImprime "@@Printer.FontName = Draft 5cpi"
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@@Printer.CurrentY = 0"
    BioImprime "@Printer.Print " & "+------------------------------------------------------------------------------+"
    x_string_40 = g_nome_empresa
    xLinha = "|                                                                  Página, ___ |"
    Mid(xLinha, 3, 40) = x_string_40
    Mid(xLinha, 76, 3) = Format(lPagina, "000")
    BioImprime "@Printer.Print " & xLinha
    '                   1         2         3         4         5         6         7         8
    '          12345678901234567890123456789012345678901234567890123456789012345678901234567890
    '                                              123456789012345678901234567890
    xLinha = "| RELAÇÃO DAS NOTAS DE ABASTECIMENTO POR CLIENTE            CIDADE, __/__/____ |"
    i = Len(g_cidade_empresa)
    Mid(xLinha, 37 + 30 - i, i) = g_cidade_empresa
    Mid(xLinha, 69, 10) = txtDataEmissao.Text
    BioImprime "@Printer.Print " & xLinha
    'x_string_40 = Mid(cbo_tipo_movimento, 3, Len(cbo_tipo_movimento))
    'BioImprime "@Printer.Print " & "| Tipo de Movimento.: " & x_string_40 & "                 |"
    xLinha = "| Periodo de Referência: __/__/____ a __/__/____                               |"
    Mid(xLinha, 26, 10) = txtDataInicial.Text
    Mid(xLinha, 39, 10) = txtDataFinal.Text
    If chkNotaConferida.Value = 1 Then
        Mid(xLinha, 57, 22) = "** NOTAS CONFERIDAS **"
    End If
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontBold = True"
    xLinha = "| Cliente.:                                                                    |"
    Mid(xLinha, 13, 40) = Cliente.RazaoSocial
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@@Printer.FontBold = False"
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    xLinha = "| Endereco..........:                                                    Bairro...:                                                     |"
    Mid(xLinha, 23, 40) = Cliente.Endereco
    Mid(xLinha, 85, 30) = Cliente.Bairro
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| Cidade............:                                                    Telefone.:                                                     |"
    Mid(xLinha, 23, 20) = Cliente.Cidade
    Mid(xLinha, 85, 20) = fMascaraTelefone(Cliente.Telefone)
    BioImprime "@Printer.Print " & xLinha
    xLinha = "+----------+----------+------+------------------------------------------+----------------+----------------+----------------+------------+"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| DATA  DE |  NUMERO  | COD. | DISCRIMINAÇÃO DAS MERCADORIAS            |                |   QUANTIDADE   |      VALOR     | CODIGO DA  |"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| EMISSAO  |   NOTA   | PROD |                                          |                |                |      TOTAL     | EMPRESA    |"
    If chkValorLiquido.Value = 1 Then
        Mid(xLinha, 108, 16) = "     LÍQUIDO    "
    End If
    BioImprime "@Printer.Print " & xLinha
    xLinha = "+----------+----------+------+------------------------------------------+----------------+----------------+----------------+------------+"
    BioImprime "@Printer.Print " & xLinha
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
        cmd_visualizar.SetFocus
    End If
End Sub
Private Sub chkUnificaEmpresa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_imprimir.SetFocus
    End If
End Sub
Private Sub cmd_data_Click()
    g_string = txtDataEmissao.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        txtDataInicial.Text = RetiraGString(1)
        txtDataFinal.Text = RetiraGString(2)
        cmd_visualizar.SetFocus
    Else
        txtDataEmissao.Text = RetiraGString(1)
        txtDataInicial.SetFocus
    End If
    g_string = ""
End Sub
Private Sub cmd_data_f_Click()
    g_string = txtDataFinal.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        txtDataInicial.Text = RetiraGString(1)
        txtDataFinal.Text = RetiraGString(2)
    Else
        txtDataFinal.Text = RetiraGString(1)
    End If
    g_string = ""
    cmd_visualizar.SetFocus
End Sub
Private Sub cmd_data_i_Click()
    g_string = txtDataInicial.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        txtDataInicial.Text = RetiraGString(1)
        txtDataFinal.Text = RetiraGString(2)
        cmd_visualizar.SetFocus
    Else
        txtDataInicial.Text = RetiraGString(1)
        txtDataFinal.SetFocus
    End If
    g_string = ""
End Sub
Private Sub cmd_imprimir_Click()
    lLocal = 1
    If ValidaCampos Then
        If SelecionaImpressoraEpson(Me) Then
            Call GravaAuditoria(1, Me.name, 7, "")
            Relatorio
        End If
    End If
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If Not IsDate(txtDataEmissao.Text) Then
        MsgBox "Informe a data de emissão.", vbInformation, "Atenção!"
        txtDataEmissao.SetFocus
    ElseIf Val(txt_cliente.Text) = 0 Then
        MsgBox "Selecione um cliente.", vbInformation, "Atenção!"
        txt_cliente.SetFocus
    ElseIf Not IsDate(txtDataInicial.Text) Then
        MsgBox "Informe a data inicial.", vbInformation, "Atenção!"
        txtDataInicial.SetFocus
    ElseIf Not IsDate(txtDataFinal.Text) Then
        MsgBox "Informe a data final.", vbInformation, "Atenção!"
        txtDataFinal.SetFocus
    ElseIf CDate(txtDataFinal.Text) < CDate(txtDataInicial.Text) Then
        MsgBox "Data final deve ser maior que a data inicial.", vbInformation, "Atenção!"
        txtDataFinal.SetFocus
    ElseIf cboGrupo.ListIndex = -1 Then
        MsgBox "Selecione um grupo.", vbInformation, "Atenção!"
        cboGrupo.SetFocus
    ElseIf cboProduto.ListIndex = -1 Then
        MsgBox "Selecione um produto.", vbInformation, "Atenção!"
        cboProduto.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_sair_Click()
    Finaliza
    Unload Me
End Sub
Private Sub cmd_visualizar_Click()
    lLocal = 0
    If ValidaCampos Then
        If SelecionaImpressoraEpson(Me) Then
            Call GravaAuditoria(1, Me.name, 6, "")
            Relatorio
        End If
    End If
End Sub
Private Sub dtcboCliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtDataInicial.SetFocus
    End If
End Sub
Private Sub dtcboCliente_LostFocus()
    If dtcboCliente.BoundText <> "" Then
        If Cliente.LocalizarCodigo(CLng(dtcboCliente.BoundText)) Then
            If Cliente.Codigo <> Val(txt_cliente.Text) Then
                txt_cliente.Text = Cliente.Codigo
            End If
            g_string = CalculaDataAbastecimentoVencimento(Cliente.CodigoVencimento, CDate(txtDataEmissao.Text))
            If g_string <> "" Then
                txtDataInicial.Text = RetiraGString(1)
                txtDataFinal.Text = RetiraGString(2)
                'l_data_vencimento = RetiraGString(3)
            End If
            g_string = ""
            txtDataInicial.SetFocus
        End If
    End If
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    If Not IsDate(txtDataEmissao.Text) Then
        txtDataEmissao.Text = Format(g_data_def, "dd/mm/yyyy")
        txtDataInicial.Text = Format(g_data_def - 1, "dd/mm/yyyy")
        txtDataFinal.Text = Format(g_data_def - 1, "dd/mm/yyyy")
        cboGrupo.ListIndex = 0
        cboProduto.ListIndex = 0
        txt_cliente.SetFocus
    End If
    Screen.MousePointer = 1
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        cmd_imprimir_Click
    ElseIf KeyCode = vbKeyF9 Then
        KeyCode = 0
        cmd_visualizar_Click
    End If
End Sub
Private Sub Form_Load()
    CentraForm Me
    Set adodcCliente.Recordset = Conectar.RsConexao("SELECT Codigo, [Razao Social] FROM Cliente WHERE Inativo = " & preparaBooleano(False) & " ORDER BY [Razao Social]")
    PreencheCboGrupo
    PreencheCboProduto
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
        If Cliente.LocalizarCodigo(CLng(txt_cliente.Text)) Then
            If Cliente.Inativo = True Then
                MsgBox "O cliente " & Trim(Cliente.RazaoSocial) & " está inativo.", vbInformation, "Atenção!"
                dtcboCliente.BoundText = ""
                txt_cliente.SetFocus
                Exit Sub
            Else
                dtcboCliente.BoundText = CLng(txt_cliente.Text)
                txtDataInicial.SetFocus
                Exit Sub
            End If
        Else
            MsgBox "Cliente não cadastrado.", vbInformation, "Atenção!"
            txt_cliente.SetFocus
            Exit Sub
        End If
    End If
End Sub
Private Sub txtDataEmissao_GotFocus()
    txtDataEmissao.Text = fDesmascaraData(txtDataEmissao.Text)
    txtDataEmissao.SelStart = 0
    txtDataEmissao.SelLength = 4
    txtDataEmissao.MaxLength = 8
End Sub
Private Sub txtDataEmissao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtDataInicial.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txtDataEmissao_LostFocus()
    txtDataEmissao.MaxLength = 10
    txtDataEmissao.Text = fMascaraData(txtDataEmissao.Text)
End Sub
Private Sub txtDataFinal_GotFocus()
    txtDataFinal.Text = fDesmascaraData(txtDataFinal.Text)
    txtDataFinal.SelStart = 0
    txtDataFinal.SelLength = 4
    txtDataFinal.MaxLength = 8
End Sub
Private Sub txtDataFinal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        chkUnificaEmpresa.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txtDataFinal_LostFocus()
    txtDataFinal.MaxLength = 10
    txtDataFinal.Text = fMascaraData(txtDataFinal.Text)
End Sub
Private Sub txtDataInicial_GotFocus()
    txtDataInicial.Text = fDesmascaraData(txtDataInicial.Text)
    txtDataInicial.SelStart = 0
    txtDataInicial.SelLength = 4
    txtDataInicial.MaxLength = 8
End Sub
Private Sub txtDataInicial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtDataFinal.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txtDataInicial_LostFocus()
    txtDataInicial.MaxLength = 10
    txtDataInicial.Text = fMascaraData(txtDataInicial.Text)
    If IsDate(txtDataInicial.Text) Then
        txtDataFinal.Text = txtDataInicial.Text
    End If
End Sub
