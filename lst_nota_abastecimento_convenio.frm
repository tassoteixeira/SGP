VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form lst_nota_abastecimento_convenio 
   Caption         =   "Emissão das Notas de Abastecimento por Convênio"
   ClientHeight    =   3195
   ClientLeft      =   3990
   ClientTop       =   2010
   ClientWidth     =   6795
   Icon            =   "lst_nota_abastecimento_convenio.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "lst_nota_abastecimento_convenio.frx":030A
   ScaleHeight     =   3195
   ScaleWidth      =   6795
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1140
      Picture         =   "lst_nota_abastecimento_convenio.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Visualiza notas de abastecimento por convênio."
      Top             =   2220
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3000
      Picture         =   "lst_nota_abastecimento_convenio.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Imprime notas de abastecimento por convênio."
      Top             =   2220
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4860
      Picture         =   "lst_nota_abastecimento_convenio.frx":3074
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   2220
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   1995
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6555
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   5940
         Picture         =   "lst_nota_abastecimento_convenio.frx":4706
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2760
         Picture         =   "lst_nota_abastecimento_convenio.frx":59E0
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2760
         Picture         =   "lst_nota_abastecimento_convenio.frx":6CBA
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.ComboBox cbo_cliente_conveniado 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1500
         Width           =   4755
      End
      Begin MSMask.MaskEdBox msk_data_f 
         Height          =   315
         Left            =   4860
         TabIndex        =   8
         Top             =   660
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_data_i 
         Height          =   315
         Left            =   1680
         TabIndex        =   5
         Top             =   660
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_data 
         Height          =   315
         Left            =   1680
         TabIndex        =   2
         Top             =   240
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSAdodcLib.Adodc adodcCliente 
         Height          =   330
         Left            =   3180
         Top             =   1080
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
         Bindings        =   "lst_nota_abastecimento_convenio.frx":7F94
         Height          =   315
         Left            =   1680
         TabIndex        =   11
         Top             =   1080
         Width           =   4755
         _ExtentX        =   8387
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Razao Social"
         BoundColumn     =   "Codigo"
         Text            =   "dtcboCliente"
      End
      Begin VB.Label Label3 
         Caption         =   "Clie&nte conveniado"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   1500
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "C&liente/Convênio"
         Height          =   315
         Index           =   8
         Left            =   120
         TabIndex        =   10
         Top             =   1080
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
         Left            =   3840
         TabIndex        =   7
         Top             =   660
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "&Data inicial"
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   660
         Width           =   1515
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   2460
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "lst_nota_abastecimento_convenio"
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
Dim l_conveniado As Long
Dim lSubTotal As Currency
Dim lTotal As Currency
Dim rstClienteConveniado As adodb.Recordset
Dim rstMovNotaAbastecimento As adodb.Recordset
Dim lSQl As String

Private Cliente As New cCliente
Private ClienteConveniado As New cClienteConveniado
Private Produto As New cProduto
Function BuscaDatas() As Boolean
    BuscaDatas = False
    msk_data_i.Text = "__/__/____"
    msk_data_f.Text = "__/__/____"
'    With tbl_movimento_nota
'        .Index = "id_cliente_data"
'        If .RecordCount > 0 Then
'            .Seek ">", Val(dtcboCliente.BoundText), CDate("01/01/1900"), 0, 0, 0, 0
'            If Not .NoMatch Then
'                If ![Codigo do Cliente] = Val(dtcboCliente.BoundText) Then
'                    msk_data_i = ![Data do Abastecimento]
'                    .Seek "<", Val(dtcboCliente.BoundText), CDate("31/12/2500"), 0, 0, 0, 0
'                    If Not .NoMatch Then
'                        If ![Codigo do Cliente] = Val(dtcboCliente.BoundText) Then
'                            msk_data_f = ![Data do Abastecimento]
'                        End If
'                    End If
'                End If
'            End If
'        End If
'        .Index = "id_conveniado"
'    End With
End Function
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set Cliente = Nothing
    Set ClienteConveniado = Nothing
    Set Produto = Nothing
End Sub
Private Sub PreencheCboClienteConveniado()
    cbo_cliente_conveniado.Clear
    cbo_cliente_conveniado.AddItem "Todos os Clientes"
    cbo_cliente_conveniado.ItemData(cbo_cliente_conveniado.NewIndex) = 0
    
    lSQl = ""
    lSQl = lSQl & "SELECT [Codigo do Conveniado], Nome"
    lSQl = lSQl & "  FROM Cliente_Conveniado"
    lSQl = lSQl & " WHERE [Codigo do Convenio] = " & Cliente.CodigoConvenio
    lSQl = lSQl & " ORDER BY Nome"
    Set rstClienteConveniado = Conectar.RsConexao(lSQl)
    
    If rstClienteConveniado.RecordCount > 0 Then
        Do Until rstClienteConveniado.EOF
            cbo_cliente_conveniado.AddItem rstClienteConveniado!Nome
            cbo_cliente_conveniado.ItemData(cbo_cliente_conveniado.NewIndex) = rstClienteConveniado![Codigo do Conveniado]
            rstClienteConveniado.MoveNext
        Loop
    End If
    rstClienteConveniado.Close
    Set rstClienteConveniado = Nothing
End Sub
Private Sub ZeraVariaveis()
    lLinha = 0
    lPagina = 0
    lSubTotal = 0
    lTotal = 0
    l_conveniado = 0
End Sub
Private Sub Relatorio()
    ZeraVariaveis
    'Verifica movimento
    
    lSQl = ""
    lSQl = lSQl & "SELECT TOP 2 [Data do Abastecimento]"
    lSQl = lSQl & "  FROM Movimento_Nota_Abastecimento"
    lSQl = lSQl & " WHERE [Codigo do Cliente] = " & Cliente.Codigo
    If CLng(cbo_cliente_conveniado.ItemData(cbo_cliente_conveniado.ListIndex)) > 0 Then
        lSQl = lSQl & "   AND [Codigo do Conveniado] = " & CLng(cbo_cliente_conveniado.ItemData(cbo_cliente_conveniado.ListIndex))
    Else
        lSQl = lSQl & "   AND [Codigo do Conveniado] > 0"
    End If
    lSQl = lSQl & "   AND [Data do Abastecimento] >= " & preparaData(CDate(msk_data_i.Text))
    lSQl = lSQl & "   AND [Data do Abastecimento] <= " & preparaData(CDate(msk_data_f.Text))
    If CLng(cbo_cliente_conveniado.ItemData(cbo_cliente_conveniado.ListIndex)) > 0 Then
        lSQl = lSQl & "   AND [Codigo do Conveniado] = " & CLng(cbo_cliente_conveniado.ItemData(cbo_cliente_conveniado.ListIndex))
    End If
    lSQl = lSQl & " ORDER BY [Data do Abastecimento]"
    Set rstMovNotaAbastecimento = Conectar.RsConexao(lSQl)
    If rstMovNotaAbastecimento.RecordCount > 0 Then
        rstMovNotaAbastecimento.Close
        ImpDados
    End If
    Set rstMovNotaAbastecimento = Nothing
    cmd_sair.SetFocus
End Sub
Private Sub ImpDados()
    Dim x_linha As String
    
    'loop movimento de notas de abastecimento
    lSQl = ""
    lSQl = lSQl & "SELECT [Codigo do Conveniado], Nome"
    lSQl = lSQl & "  FROM Cliente_Conveniado"
    lSQl = lSQl & " WHERE [Codigo do Convenio] = " & Cliente.CodigoConvenio
    If CLng(cbo_cliente_conveniado.ItemData(cbo_cliente_conveniado.ListIndex)) > 0 Then
        lSQl = lSQl & "   AND [Codigo do Conveniado] = " & CLng(cbo_cliente_conveniado.ItemData(cbo_cliente_conveniado.ListIndex))
    End If
    lSQl = lSQl & " ORDER BY Nome"
    Set rstClienteConveniado = Conectar.RsConexao(lSQl)
    
    If rstClienteConveniado.RecordCount > 0 Then
        Do Until rstClienteConveniado.EOF
            lSQl = ""
            lSQl = lSQl & "SELECT [Data do Abastecimento], [Numero da Nota], [Valor Unitario], Quantidade, [Valor Total], [Codigo do Produto2], [Codigo do Conveniado], Empresa"
            lSQl = lSQl & "  FROM Movimento_Nota_Abastecimento"
            lSQl = lSQl & " WHERE [Codigo do Cliente] = " & Cliente.Codigo
            lSQl = lSQl & "   AND [Codigo do Conveniado] = " & rstClienteConveniado![Codigo do Conveniado]
            lSQl = lSQl & "   AND [Data do Abastecimento] >= " & preparaData(CDate(msk_data_i.Text))
            lSQl = lSQl & "   AND [Data do Abastecimento] <= " & preparaData(CDate(msk_data_f.Text))
            If CLng(cbo_cliente_conveniado.ItemData(cbo_cliente_conveniado.ListIndex)) > 0 Then
                lSQl = lSQl & "   AND [Codigo do Conveniado] = " & CLng(cbo_cliente_conveniado.ItemData(cbo_cliente_conveniado.ListIndex))
            End If
            lSQl = lSQl & " ORDER BY [Data do Abastecimento], [Numero da Nota], [Codigo do Produto2]"
            Set rstMovNotaAbastecimento = Conectar.RsConexao(lSQl)
            Do Until rstMovNotaAbastecimento.EOF
                If lPagina = 0 Then
                    ImpCab
                End If
                If lLinha >= 57 Then
                    x_linha = "+------------+----------+--------+------------------------------------------+----------+----------------+-------------------------------+"
                    BioImprime "@Printer.Print " & x_linha
                    BioImprime "@@Printer.NewPage"
                    ImpCab
                End If
                If rstMovNotaAbastecimento![Codigo do Conveniado] <> l_conveniado Then
                    ImpClienteConveniado
                End If
                ImpProduto
                lSubTotal = lSubTotal + rstMovNotaAbastecimento![Valor Total]
                lTotal = lTotal + rstMovNotaAbastecimento![Valor Total]
                rstMovNotaAbastecimento.MoveNext
            Loop
            rstMovNotaAbastecimento.Close
            rstClienteConveniado.MoveNext
        Loop
    End If
    rstClienteConveniado.Close
    Set rstClienteConveniado = Nothing
    
    
'    With tbl_movimento_nota
'        If tbl_cliente_conveniado.RecordCount > 0 Then
'            tbl_cliente_conveniado.Index = "id_nome"
'            tbl_cliente_conveniado.Seek ">", " ", 0, 0
'            If Not tbl_cliente_conveniado.NoMatch Then
'                Do Until tbl_cliente_conveniado.EOF
'                    If tbl_cliente_conveniado![Codigo do Convenio] = tbl_cliente![Codigo do Convenio] Then
'                        .Seek ">", CLng(dtcboCliente.BoundText), tbl_cliente_conveniado![Codigo do Conveniado], CDate(msk_data_i), 0, 0, 0, 0
'                        If Not .NoMatch Then
'                            Do Until .EOF
'                                If ![Codigo do Conveniado] <> tbl_cliente_conveniado![Codigo do Conveniado] Then
'                                    Exit Do
'                                End If
'                                If ![Data do Abastecimento] >= CDate(msk_data_i) And ![Data do Abastecimento] <= CDate(msk_data_f) Then
'                                    If ![Codigo do Conveniado] = CLng(cbo_cliente_conveniado.ItemData(cbo_cliente_conveniado.ListIndex)) Or CLng(cbo_cliente_conveniado.ItemData(cbo_cliente_conveniado.ListIndex)) = 0 Then
'                                        If lPagina = 0 Then
'                                            ImpCab
'                                        End If
'                                        If lLinha >= 57 Then
'                                            x_linha = "+------------+----------+--------+------------------------------------------+----------+----------------+-------------------------------+"
'                                            BioImprime "@Printer.Print " & x_linha
'                                            BioImprime "@@Printer.NewPage"
'                                            ImpCab
'                                        End If
'                                        If ![Codigo do Conveniado] <> l_conveniado Then
'                                            ImpClienteConveniado
'                                        End If
'                                        ImpProduto
'                                        lSubTotal = lSubTotal + ![Valor Total]
'                                        lTotal = lTotal + ![Valor Total]
'                                    End If
'                                End If
'                                .MoveNext
'                            Loop
'                        End If
'                    End If
'                    tbl_cliente_conveniado.MoveNext
'                Loop
'            End If
'        End If
'    End With
    If lTotal > 0 Then
        ImpSubTotal
        ImpTotal
        BioImprime "@@Printer.EndDoc"
        BioFechaImprime
        g_string = lLocal & lNomeArquivo & "|@|Relatório de Nota de Abastecimento por Convênio|@|"
        frm_preview.Show 1
    End If
End Sub
Private Sub ImpClienteConveniado()
    Dim x_linha As String * 80
    Dim i As Integer
    If lSubTotal > 0 Then
        ImpSubTotal
    End If
    x_linha = Space(80)
    Mid(x_linha, 1, 15) = "| Conveniado.: "
    i = Len(Format(rstMovNotaAbastecimento![Codigo do Conveniado], "###,###"))
    Mid(x_linha, 17 + 7 - i, i) = Format(rstMovNotaAbastecimento![Codigo do Conveniado], "###,###")
    Mid(x_linha, 25, 40) = rstClienteConveniado!Nome
    Mid(x_linha, 80, 1) = "|"
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & x_linha
    l_conveniado = rstMovNotaAbastecimento![Codigo do Conveniado]
    lLinha = lLinha + 1
End Sub
Private Sub ImpProduto()
    Dim x_linha As String * 137
    Dim x_nome_produto As String
    Dim i As Integer
    x_linha = Space(137)
    Mid(x_linha, 1, 1) = "|"
    Mid(x_linha, 3, 10) = Format(rstMovNotaAbastecimento![Data do Abastecimento], "dd/mm/yyyy")
    Mid(x_linha, 14, 1) = "|"
    i = Len(Format(rstMovNotaAbastecimento![Numero da Nota], "###,##0"))
    Mid(x_linha, 17 + 7 - i, i) = Format(rstMovNotaAbastecimento![Numero da Nota], "###,##0")
    Mid(x_linha, 25, 1) = "|"
    If Produto.LocalizarCodigo(rstMovNotaAbastecimento![Codigo do Produto2]) Then
        x_nome_produto = Produto.Nome
    Else
        x_nome_produto = "** Não Cadastrado **"
    End If
    i = Len(Format(rstMovNotaAbastecimento![Codigo do Produto2], "##,000"))
    Mid(x_linha, 27 + 6 - i, i) = Format(rstMovNotaAbastecimento![Codigo do Produto2], "##,000")
    Mid(x_linha, 34, 1) = "|"
    Mid(x_linha, 36, 40) = x_nome_produto
    Mid(x_linha, 77, 1) = "|"
    i = Len(Format(rstMovNotaAbastecimento!Quantidade, "####,##0.00"))
    Mid(x_linha, 76 + 11 - i, i) = Format(rstMovNotaAbastecimento!Quantidade, "####,##0.00")
    Mid(x_linha, 88, 1) = "|"
    i = Len(Format(rstMovNotaAbastecimento![Valor Total], "###,###,##0.00"))
    Mid(x_linha, 90 + 14 - i, i) = Format(rstMovNotaAbastecimento![Valor Total], "###,###,##0.00")
    Mid(x_linha, 105, 1) = "|"
    Mid(x_linha, 120, 2) = Format(rstMovNotaAbastecimento!Empresa, "00")
    Mid(x_linha, 137, 1) = "|"
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@Printer.Print " & x_linha
    lLinha = lLinha + 1
End Sub
Private Sub ImpSubTotal()
    Dim x_linha As String * 137
    Dim i As Integer
    x_linha = Space(137)
    Mid(x_linha, 1, 1) = "|"
    Mid(x_linha, 68, 22) = "Total do Conveniado.: "
    i = Len(Format(lSubTotal, "###,###,##0.00"))
    Mid(x_linha, 90 + 14 - i, i) = Format(lSubTotal, "###,###,##0.00")
    Mid(x_linha, 105, 1) = "|"
    Mid(x_linha, 137, 1) = "|"
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@@Printer.FontBold = True"
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontBold = False"
    BioImprime "@Printer.Print " & "+------------+----------+--------+------------------------------------------+----------+----------------+-------------------------------+"
    lLinha = lLinha + 2
    lSubTotal = 0
End Sub
Private Sub ImpTotal()
    Dim x_linha As String * 137
    Dim i As Integer
    If lLinha > 47 Then
        x_linha = " "
        BioImprime "@Printer.Print " & x_linha
        BioImprime "@@Printer.NewPage"
        ImpCab
    End If
    x_linha = Space(137)
    Mid(x_linha, 1, 1) = "|"
    Mid(x_linha, 68, 22) = "Total Geral.........: "
    i = Len(Format(lTotal, "###,###,##0.00"))
    Mid(x_linha, 90 + 14 - i, i) = Format(lTotal, "###,###,##0.00")
    Mid(x_linha, 105, 1) = "|"
    Mid(x_linha, 137, 1) = "|"
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@@Printer.FontBold = True"
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontBold = False"
    x_linha = "+--------------------------------------------------------------------------------------+----------------+-------------------------------+"
    BioImprime "@Printer.Print " & x_linha
    x_linha = String(137, "*")
    Mid(x_linha, 1, 3) = "| ("
    i = Len(FazExtenso(lTotal))
    Mid(x_linha, 4, i) = FazExtenso(lTotal)
    Mid(x_linha, 135, 3) = ") |"
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontBold = True"
    x_linha = "|                         FAVOR ENVIAR ORDEM DE PAGAMENTO PELO BANCO SAFRA AGENCIA 3600 C/C 018523-1 GOIANIA-GO.                        |"
    x_linha = "|                                                                                                                                       |"
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontBold = False"
    x_linha = "|                                                                                                                                       |"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "|                                                             Atenciosamente,                                                           |"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "|                                                                                                                                       |"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "|                                                ________________________________________                                               |"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "|                                                                                                                                       |"
    i = Len(Trim(g_nome_empresa))
    Mid(x_linha, 50 + (40 - i) / 2, i) = Trim(g_nome_empresa)
    BioImprime "@Printer.Print " & x_linha
    x_linha = "|                                                                                                                                       |"
    BioImprime "@Printer.Print " & x_linha
    x_linha = "+---------------------------------------------------------------------------------------------------------------------------------------+"
    Mid(x_linha, 5, 22) = " Cerrado Informática. "
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & "  "
End Sub
Private Sub ImpCab()
    Dim x_string_40 As String * 40
    Dim x_linha As String
    Dim i As Integer
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
    BioImprime "@@Printer.FontBold = True"
    x_linha = "|                                                                  Página: ___ |"
    Mid(x_linha, 3, 40) = g_nome_empresa
    Mid(x_linha, 76, 3) = Format(lPagina, "000")
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontBold = False"
    '                   1         2         3         4         5         6         7         8         9        10        11        12        13     13
    '          12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567
    x_linha = "| RELAÇÃO DAS NOTAS DE ABASTECIMENTO POR CONVENIO           CIDADE, __/__/____ |"
    i = Len(g_cidade_empresa)
    Mid(x_linha, 37 + 30 - i, i) = g_cidade_empresa
    Mid(x_linha, 69, 10) = msk_data.Text
    BioImprime "@Printer.Print " & x_linha
    x_linha = "| NOME DO CONVÊNIO........:                                                    |"
    Mid(x_linha, 29, 3) = Format(CLng(dtcboCliente.BoundText), "000")
    Mid(x_linha, 33, 40) = dtcboCliente.Text
    BioImprime "@Printer.Print " & x_linha
    x_linha = "| REFERENTE AO PERÍODO DE.: __/__/____ A __/__/____                            |"
    Mid(x_linha, 29, 10) = Format(msk_data_i, "dd/mm/yyyy")
    Mid(x_linha, 42, 10) = Format(msk_data_f, "dd/mm/yyyy")
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@@Printer.FontBold = False"
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@Printer.Print " & "+------------+----------+--------+------------------------------------------+----------+----------------+-------------------------------+"
    BioImprime "@Printer.Print " & "|    DATA    | N.  NOTA | CÓDIGO | DISCRIMINAÇÃO DOS PRODUTOS               |QUANTIDADE|VALOR DO PRODUTO|            EMPRESA            |"
    BioImprime "@Printer.Print " & "+------------+----------+--------+------------------------------------------+----------+----------------+-------------------------------+"
End Sub
Private Sub cbo_cliente_conveniado_GotFocus()
    SendMessageLong cbo_cliente_conveniado.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_cliente_conveniado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_imprimir.SetFocus
    End If
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i = RetiraGString(1)
        msk_data_f = RetiraGString(2)
        dtcboCliente.SetFocus
    Else
        msk_data = RetiraGString(1)
        msk_data_i.SetFocus
    End If
    g_string = " "
End Sub
Private Sub cmd_data_f_Click()
    g_string = msk_data_f
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i = RetiraGString(1)
        msk_data_f = RetiraGString(2)
    Else
        msk_data_f = RetiraGString(1)
    End If
    g_string = " "
    dtcboCliente.SetFocus
End Sub
Private Sub cmd_data_i_Click()
    g_string = msk_data_i
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i = RetiraGString(1)
        msk_data_f = RetiraGString(2)
        dtcboCliente.SetFocus
    Else
        msk_data_i = RetiraGString(1)
        msk_data_f.SetFocus
    End If
    g_string = " "
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
    If Not IsDate(msk_data.Text) Then
        MsgBox "Informe a data de emissão.", vbInformation, "Atenção!"
        msk_data.SetFocus
    ElseIf Not IsDate(msk_data_i.Text) Then
        MsgBox "Informe a data inicial.", vbInformation, "Atenção!"
        msk_data_i.SetFocus
    ElseIf Not IsDate(msk_data_f.Text) Then
        MsgBox "Informe a data final.", vbInformation, "Atenção!"
        msk_data_f.SetFocus
    ElseIf CDate(msk_data_f.Text) < CDate(msk_data_i.Text) Then
        MsgBox "Data final deve ser maior que a data inicial.", vbInformation, "Atenção!"
        msk_data_f.SetFocus
    ElseIf Not dtcboCliente.BoundText <> "" Then
        MsgBox "Escolha o cliente/convênio.", vbInformation, "Atenção!"
        dtcboCliente.SetFocus
    ElseIf cbo_cliente_conveniado.ListIndex = -1 Then
        MsgBox "Escolha o cliente conveniado.", vbInformation, "Atenção!"
        cbo_cliente_conveniado.SetFocus
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
        cbo_cliente_conveniado.SetFocus
    End If
End Sub
Private Sub dtcboCliente_LostFocus()
    If dtcboCliente.BoundText <> "" Then
        If Cliente.LocalizarCodigo(CLng(dtcboCliente.BoundText)) Then
            PreencheCboClienteConveniado
        End If
    End If
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    If Not IsDate(msk_data.Text) Then
        msk_data.Text = Format(g_data_def, "dd/mm/yyyy")
        dtcboCliente.BoundText = 26
        PreencheCboClienteConveniado
        cbo_cliente_conveniado.ListIndex = 0
        BuscaDatas
        cmd_imprimir.SetFocus
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
    
    Set adodcCliente.Recordset = Conectar.RsConexao("SELECT Codigo, [Razao Social] FROM Cliente WHERE [Codigo do Convenio] > 1 ORDER BY [Razao Social]")
End Sub
Private Sub msk_data_f_GotFocus()
    msk_data_f.SelStart = 0
    msk_data_f.SelLength = 2
End Sub
Private Sub msk_data_f_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dtcboCliente.SetFocus
    End If
End Sub
Private Sub msk_data_i_GotFocus()
    msk_data_i.SelStart = 0
    msk_data_i.SelLength = 2
End Sub
Private Sub msk_data_i_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data_f.SetFocus
    End If
End Sub
Private Sub msk_data_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data_i.SetFocus
    End If
End Sub
