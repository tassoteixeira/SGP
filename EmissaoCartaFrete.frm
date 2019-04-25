VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form EmissaoCartaFrete 
   Caption         =   "Emissão da Carta Frete"
   ClientHeight    =   2700
   ClientLeft      =   3990
   ClientTop       =   2010
   ClientWidth     =   7545
   Icon            =   "EmissaoCartaFrete.frx":0000
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   Picture         =   "EmissaoCartaFrete.frx":030A
   ScaleHeight     =   2700
   ScaleWidth      =   7545
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1140
      Picture         =   "EmissaoCartaFrete.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Visualiza a carta frete."
      Top             =   1740
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3000
      Picture         =   "EmissaoCartaFrete.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Imprime a carta frete."
      Top             =   1740
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4860
      Picture         =   "EmissaoCartaFrete.frx":3074
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   1740
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   1515
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7275
      Begin VB.TextBox txt_cliente 
         Height          =   285
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   11
         Top             =   1080
         Width           =   795
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2760
         Picture         =   "EmissaoCartaFrete.frx":4706
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2760
         Picture         =   "EmissaoCartaFrete.frx":59E0
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   6660
         Picture         =   "EmissaoCartaFrete.frx":6CBA
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin MSMask.MaskEdBox msk_data_f 
         Height          =   315
         Left            =   5580
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
         Left            =   4020
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
         Bindings        =   "EmissaoCartaFrete.frx":7F94
         Height          =   315
         Left            =   2520
         TabIndex        =   12
         Top             =   1080
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
         Caption         =   "C&liente"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "&Data inicial"
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   660
         Width           =   1515
      End
      Begin VB.Label Label2 
         Caption         =   "D&ata final"
         Height          =   315
         Left            =   4560
         TabIndex        =   7
         Top             =   660
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Data de &emissão"
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1515
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   1980
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "EmissaoCartaFrete"
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
Dim lTotalAbastecimento As Currency
Dim lTotalTrocoCheque As Currency
Dim lTotalTrocoDinheiro As Currency
Dim lTotal As Currency
Dim lSQL As String

Private Cliente As New cCliente
Private rsMovCartaFrete As New adodb.Recordset
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set Cliente = Nothing
    Set rsMovCartaFrete = Nothing
End Sub
Private Sub ZeraVariaveis()
    lLinha = 0
    lPagina = 0
    lTotalAbastecimento = 0
    lTotalTrocoCheque = 0
    lTotalTrocoDinheiro = 0
    lTotal = 0
End Sub
Private Sub Relatorio()
    ZeraVariaveis
    Dim xImprimeClienteEspecifico As Boolean
    
    xImprimeClienteEspecifico = False
    If Len(txt_cliente.Text) > 0 Then
        If CInt(txt_cliente.Text) > 0 Then
            xImprimeClienteEspecifico = True
        End If
    End If
    
    
    If xImprimeClienteEspecifico = True Then
        'Prepara SQL MovimentoCartaFrete
        lSQL = "SELECT Data, Numero, Nome, Veiculo, [Valor da Carta], "
        lSQL = lSQL & "[Valor do Abastecimento], [Troco em Dinheiro Pista], [Troco em Cheque], [Numero da Conta do Cheque]"
        lSQL = lSQL & " FROM MovimentoCartaFrete"
        lSQL = lSQL & " WHERE Empresa = " & g_empresa
        lSQL = lSQL & "   AND [Codigo do Cliente] = " & CInt(txt_cliente.Text)
        lSQL = lSQL & "   AND Data >= " & preparaData(CDate(msk_data_i.Text))
        lSQL = lSQL & "   AND Data <= " & preparaData(CDate(msk_data_f.Text))
        lSQL = lSQL & " ORDER BY Data, Numero"
        Set rsMovCartaFrete = New adodb.Recordset
        Set rsMovCartaFrete = Conectar.RsConexao(lSQL)
        
        LoopImprimeDadosCliente
    Else
        'Prepara SQL MovimentoCartaFrete
        lSQL = "SELECT MovimentoCartaFrete.Data, MovimentoCartaFrete.Periodo, MovimentoCartaFrete.[Tipo do Movimento], MovimentoCartaFrete.[Codigo do Cliente], MovimentoCartaFrete.Numero, MovimentoCartaFrete.Nome, MovimentoCartaFrete.Veiculo, MovimentoCartaFrete.[Valor da Carta], "
        lSQL = lSQL & "MovimentoCartaFrete.[Valor do Abastecimento], MovimentoCartaFrete.[Troco em Dinheiro Pista], MovimentoCartaFrete.[Troco em Cheque], MovimentoCartaFrete.[Numero da Conta do Cheque], Cliente.[Razao Social]"
        lSQL = lSQL & " FROM MovimentoCartaFrete, Cliente"
        lSQL = lSQL & " WHERE MovimentoCartaFrete.Empresa = " & g_empresa
        lSQL = lSQL & "   AND MovimentoCartaFrete.Data >= " & preparaData(CDate(msk_data_i.Text))
        lSQL = lSQL & "   AND MovimentoCartaFrete.Data <= " & preparaData(CDate(msk_data_f.Text))
        lSQL = lSQL & "   AND Cliente.Codigo = MovimentoCartaFrete.[Codigo do Cliente]"
        lSQL = lSQL & " ORDER BY MovimentoCartaFrete.Data, MovimentoCartaFrete.Numero"
        Set rsMovCartaFrete = New adodb.Recordset
        Set rsMovCartaFrete = Conectar.RsConexao(lSQL)
        
        LoopImprimeDados
    End If

    cmd_sair.SetFocus
End Sub
Private Sub LoopImprimeDados()
    Dim xLinha As String
    
    If rsMovCartaFrete.RecordCount > 0 Then
        ImpCab
        Do Until rsMovCartaFrete.EOF
            If lLinha >= 57 Then
                xLinha = "+----------+---------+------------+----------+------------------------------------------+--------------------------------+--------------+"
                Mid(xLinha, 28, 22) = " Cerrado Informática. "
                BioImprime "@Printer.Print " & xLinha
                BioImprime "@@Printer.NewPage"
                ImpCab
            End If
            Call ImpDet
            rsMovCartaFrete.MoveNext
        Loop
        ImpTotal
        BioImprime "@@Printer.EndDoc"
        BioFechaImprime
        g_string = lLocal & lNomeArquivo & "|@|Relatório de Carta Frete|@|"
        frm_preview.Show 1
    Else
        MsgBox "Não existe Carta-Frete no período informado!", vbInformation, "Atenção!"
    End If
    rsMovCartaFrete.Close
End Sub
Private Sub LoopImprimeDadosCliente()
    Dim xLinha As String
    
    If rsMovCartaFrete.RecordCount > 0 Then
        ImpCabCliente
        Do Until rsMovCartaFrete.EOF
            If lLinha >= 57 Then
                xLinha = "+----------+------+------+---------------------------------------------------------------------+----------+---------------+-------------+"
                Mid(xLinha, 28, 22) = " Cerrado Informática. "
                BioImprime "@Printer.Print " & xLinha
                BioImprime "@@Printer.NewPage"
                ImpCabCliente
            End If
            Call ImpDetCliente
            rsMovCartaFrete.MoveNext
        Loop
        ImpTotalCliente
        BioImprime "@@Printer.EndDoc"
        BioFechaImprime
        g_string = lLocal & lNomeArquivo & "|@|Relatório de Carta Frete|@|"
        frm_preview.Show 1
    Else
        MsgBox "Não existe Carta-Frete deste cliente no período informado!", vbInformation, "Atenção!"
    End If
    rsMovCartaFrete.Close
End Sub
Private Sub ImpDet()
    Dim xLinha As String
    Dim i As Integer
    '                 10        20        30        40        50        60        70        80        90       100       110       120       130
    '         12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567
    xLinha = "|            |            |          |                                          |                                        |              |"
    Mid(xLinha, 3, 10) = Format(rsMovCartaFrete("Data").Value, "dd/mm/yyyy")
    Mid(xLinha, 21, 1) = rsMovCartaFrete("Periodo").Value
    i = Len(Format(rsMovCartaFrete("Numero").Value, "#######0"))
    Mid(xLinha, 29 + 8 - i, i) = Format(rsMovCartaFrete("Numero").Value, "#######0")
    Mid(xLinha, 40, 40) = rsMovCartaFrete("Razao Social").Value
    Mid(xLinha, 83, 38) = rsMovCartaFrete("Nome").Value
    i = Len(Format(rsMovCartaFrete("Valor da Carta").Value, "#####,##0.00"))
    Mid(xLinha, 124 + 12 - i, i) = Format(rsMovCartaFrete("Valor da Carta").Value, "#####,##0.00")
    BioImprime "@Printer.Print " & xLinha
    
    '''               10        20        30        40        50        60        70        80        90       100       110       120       130
    '''       12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567
    xLinha = "|            |            |          |              |              |            |                                        |              |"
    i = Len(Format(rsMovCartaFrete("Troco em Dinheiro Pista").Value, "####,##0.00"))
    Mid(xLinha, 2 + 11 - i, i) = Format(rsMovCartaFrete("Troco em Dinheiro Pista").Value, "####,##0.00")
    Mid(xLinha, 21, 1) = rsMovCartaFrete("Tipo do Movimento").Value
    If rsMovCartaFrete("Numero da Conta do Cheque").Value <> "" Then
        i = Len(Format(rsMovCartaFrete("Numero da Conta do Cheque").Value, "#######0"))
        Mid(xLinha, 29 + 8 - i, i) = Format(rsMovCartaFrete("Numero da Conta do Cheque").Value, "#######0")
    End If
    If rsMovCartaFrete("Troco em Cheque").Value > 0 Then
        i = Len(Format(rsMovCartaFrete("Troco em Cheque").Value, "####,##0.00"))
        Mid(xLinha, 41 + 11 - i, i) = Format(rsMovCartaFrete("Troco em Cheque").Value, "####,##0.00")
    End If
    i = Len(Format(rsMovCartaFrete("Valor do Abastecimento").Value, "#####,##0.00"))
    Mid(xLinha, 55 + 12 - i, i) = Format(rsMovCartaFrete("Valor do Abastecimento").Value, "#####,##0.00")
    Mid(xLinha, 83, 30) = rsMovCartaFrete("Veiculo").Value
    BioImprime "@Printer.Print " & xLinha
    
    xLinha = "+------------+------------+----------+------------------------------------------+----------------------------------------+--------------+"
    BioImprime "@Printer.Print " & xLinha
    lTotal = lTotal + rsMovCartaFrete("Valor da Carta").Value
    lTotalAbastecimento = lTotalAbastecimento + rsMovCartaFrete("Valor do Abastecimento").Value
    lTotalTrocoCheque = lTotalTrocoCheque + rsMovCartaFrete("Troco em Cheque").Value
    lTotalTrocoDinheiro = lTotalTrocoDinheiro + rsMovCartaFrete("Troco em Dinheiro Pista").Value
    lLinha = lLinha + 3
End Sub
Private Sub ImpDetCliente()
    Dim xLinha As String
    Dim i As Integer
    '''               10        20        30        40        50        60        70        80        90       100       110       120       130
    '''       12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567
    xLinha = "         |            |            |                                          |                                |              |          "
    Mid(xLinha, 12, 10) = Format(rsMovCartaFrete("Data").Value, "dd/mm/yyyy")
    i = Len(Format(rsMovCartaFrete("Numero").Value, "#########0"))
    Mid(xLinha, 25 + 10 - i, i) = Format(rsMovCartaFrete("Numero").Value, "#########0")
    Mid(xLinha, 38, 40) = rsMovCartaFrete("Nome").Value
    Mid(xLinha, 81, 30) = rsMovCartaFrete("Veiculo").Value
    i = Len(Format(rsMovCartaFrete("Valor da Carta").Value, "#####,##0.00"))
    Mid(xLinha, 114 + 12 - i, i) = Format(rsMovCartaFrete("Valor da Carta").Value, "#####,##0.00")
    BioImprime "@Printer.Print " & xLinha
    
    xLinha = "         |            |            |            |                             |                                |              |"
    If rsMovCartaFrete("Troco em Dinheiro Pista").Value > 0 Then
        i = Len(Format(rsMovCartaFrete("Troco em Dinheiro Pista").Value, "####,##0.00"))
        Mid(xLinha, 11 + 11 - i, i) = Format(rsMovCartaFrete("Troco em Dinheiro Pista").Value, "####,##0.00")
    End If
    If rsMovCartaFrete("Troco em Cheque").Value > 0 Then
        i = Len(Format(rsMovCartaFrete("Troco em Cheque").Value, "####,##0.00"))
        Mid(xLinha, 24 + 11 - i, i) = Format(rsMovCartaFrete("Troco em Cheque").Value, "####,##0.00")
    End If
    If rsMovCartaFrete("Numero da Conta do Cheque").Value <> "" Then
        i = Len(Format(rsMovCartaFrete("Numero da Conta do Cheque").Value, "#########0"))
        Mid(xLinha, 37 + 10 - i, i) = Format(rsMovCartaFrete("Numero da Conta do Cheque").Value, "#########0")
    End If
    i = Len(Format(rsMovCartaFrete("Valor do Abastecimento").Value, "####,##0.00"))
    Mid(xLinha, 68 + 10 - i, i) = Format(rsMovCartaFrete("Valor do Abastecimento").Value, "####,##0.00")
    BioImprime "@Printer.Print " & xLinha
    
    xLinha = "         +------------+------------+------------------------------------------+--------------------------------+--------------+          "
    BioImprime "@Printer.Print " & xLinha
    lTotal = lTotal + rsMovCartaFrete("Valor da Carta").Value
    lTotalAbastecimento = lTotalAbastecimento + rsMovCartaFrete("Valor do Abastecimento").Value
    lTotalTrocoCheque = lTotalTrocoCheque + rsMovCartaFrete("Troco em Cheque").Value
    lTotalTrocoDinheiro = lTotalTrocoDinheiro + rsMovCartaFrete("Troco em Dinheiro Pista").Value
    lLinha = lLinha + 3
End Sub
Private Sub ImpTotalCliente()
    Dim xLinha As String
    Dim i As Integer
    
    'xLinha = "         +------------+------------+------------------------------------------+--------------------------------+--------------+          "
    'BioImprime "@Printer.Print " & xLinha
    xLinha = "         |            |            |                                          |                      ** TOTAL  |              |"
    i = Len(Format(lTotalTrocoDinheiro, "####,##0.00"))
    Mid(xLinha, 11 + 11 - i, i) = Format(lTotalTrocoDinheiro, "####,##0.00")
    i = Len(Format(lTotalTrocoCheque, "####,##0.00"))
    Mid(xLinha, 24 + 11 - i, i) = Format(lTotalTrocoCheque, "####,##0.00")
    i = Len(Format(lTotalAbastecimento, "####,##0.00"))
    Mid(xLinha, 68 + 10 - i, i) = Format(lTotalAbastecimento, "####,##0.00")
    i = Len(Format(lTotal, "#####,##0.00"))
    Mid(xLinha, 114 + 12 - i, i) = Format(lTotal, "#####,##0.00")
    BioImprime "@Printer.Print " & xLinha
    xLinha = "         +------------+------------+------------------------------------------+--------------------------------+--------------+          "
    BioImprime "@Printer.Print " & xLinha
    
    BioImprime "@@Printer.FontName = Draft 10cpi"
    
    
    BioImprime "@Printer.Print " & "    "
    BioImprime "@Printer.Print " & "    "
    BioImprime "@Printer.Print " & "    "
    BioImprime "@Printer.Print " & "    "
    BioImprime "@Printer.Print " & "    "
    BioImprime "@Printer.Print " & "    "
    BioImprime "@Printer.Print " & "    "
    BioImprime "@Printer.Print " & "    "
    BioImprime "@Printer.Print " & "    "
    BioImprime "@Printer.Print " & "    "
    
    xLinha = "     Atenciosamente,                                                            "
    BioImprime "@Printer.Print " & xLinha
    
    BioImprime "@Printer.Print " & "    "
    BioImprime "@Printer.Print " & "    "
    
    xLinha = "                                                                                "
    Mid(xLinha, 6, 40) = g_nome_empresa
    BioImprime "@Printer.Print " & xLinha
    
    BioImprime "@Printer.Print " & "  "
End Sub
Private Sub ImpTotal()
    Dim xLinha As String
    Dim i As Integer
    
    '''               10        20        30        40        50        60        70        80        90       100       110       120       130
    '''       12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567
    xLinha = "|            |            |          |              |              |            |                              ** TOTAL  |              |"
    i = Len(Format(lTotalTrocoDinheiro, "####,##0.00"))
    Mid(xLinha, 2 + 11 - i, i) = Format(lTotalTrocoDinheiro, "####,##0.00")
    i = Len(Format(lTotalTrocoCheque, "#####,##0.00"))
    Mid(xLinha, 40 + 12 - i, i) = Format(lTotalTrocoCheque, "#####,##0.00")
    i = Len(Format(lTotalAbastecimento, "#####,##0.00"))
    Mid(xLinha, 55 + 12 - i, i) = Format(lTotalAbastecimento, "#####,##0.00")
    i = Len(Format(lTotal, "#####,##0.00"))
    Mid(xLinha, 124 + 12 - i, i) = Format(lTotal, "#####,##0.00")
    BioImprime "@Printer.Print " & xLinha
    xLinha = "+------------+------------+----------+------------------------------------------+----------------------------------------+--------------+"
    Mid(xLinha, 83, 21) = " Cerrado Tecnologia. "
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontName = Draft 10cpi"
End Sub
Private Sub ImpCab()
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
    BioImprime "@@Printer.FontBold = True"
    x_linha = "|                                                                  Página,     |"
    Mid(x_linha, 3, 40) = g_nome_empresa
    Mid(x_linha, 76, 3) = Format(lPagina, "000")
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontBold = False"
    x_linha = "| RELAÇÃO DE CARTA FRETE                                          , __/__/____ |"
    i = Len(g_cidade_empresa)
    Mid(x_linha, 37 + 30 - i, i) = g_cidade_empresa
    Mid(x_linha, 69, 10) = msk_data
    BioImprime "@Printer.Print " & x_linha
    x_linha = "| PERÍODO DE ABASTECIMENTO: __/__/____ A __/__/____                            |"
    Mid(x_linha, 29, 10) = msk_data_i.Text
    Mid(x_linha, 42, 10) = msk_data_f.Text
    BioImprime "@Printer.Print " & x_linha
'    x_linha = "| CONTA...................:                                                    |"
'    Mid(x_linha, 29, 40) = cbo_conta
'    BioImprime "@Printer.Print " & x_linha
'    x_linha = "| FORNECEDOR..............:                                                    |"
'    Mid(x_linha, 29, 40) = cbo_fornecedor
'    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@@Printer.FontBold = False"
    BioImprime "@Printer.Print " & "+------------+------------+----------+------------------------------------------+----------------------------------------+--------------+"
    BioImprime "@Printer.Print " & "|  DATA  DA  |  PERIODO   |  NUMERO  | RAZÃO SOCIAL                             | PROPRIETARIO / MOTORISTA               | VLR DA CARTA |"
    BioImprime "@Printer.Print " & "| TROCO  R$  | TIPO  MOV. | N.CHEQUE | TROCO CHEQUE | VLR.  ABAST. |            | VEICULO                                |              |"
    BioImprime "@Printer.Print " & "+------------+------------+----------+------------------------------------------+----------------------------------------+--------------+"
End Sub
Private Sub ImpCabCliente()
    Dim i As Integer
    Dim xLinha As String
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
    BioImprime "@Printer.Print " & "    "
    BioImprime "@Printer.Print " & "    "
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & "    "
    BioImprime "@Printer.Print " & "    "
    BioImprime "@Printer.Print " & "    "
    BioImprime "@Printer.Print " & "    "
    
    xLinha = "     Cidade                                                                     "
    i = Len(g_cidade_empresa)
    Mid(xLinha, 6, 50) = Trim(g_cidade_empresa) & ", " & Day(CDate(msk_data.Text)) & " de " & Format(CDate(msk_data.Text), "mmmm") & " de " & Year(CDate(msk_data.Text)) & "."
    BioImprime "@Printer.Print " & xLinha
    
    BioImprime "@Printer.Print " & "    "
    
    xLinha = "     À                                                                          "
    BioImprime "@Printer.Print " & xLinha
    xLinha = "                                                                                "
    Mid(xLinha, 6, 40) = Cliente.RazaoSocial
    BioImprime "@Printer.Print " & xLinha
    xLinha = "     Nesta                                                                      "
    BioImprime "@Printer.Print " & xLinha
    
    BioImprime "@Printer.Print " & "    "
    BioImprime "@Printer.Print " & "    "
    BioImprime "@Printer.Print " & "    "
    
    xLinha = "     ASSUNTO:  PAGAMENTO DE CARTAS-FRETE                                        "
    BioImprime "@Printer.Print " & xLinha
    
    BioImprime "@Printer.Print " & "    "
    BioImprime "@Printer.Print " & "    "
    
    xLinha = "     Efetuamos no período de __/__/____ a __/__/____, o pagamento das           "
    Mid(xLinha, 30, 10) = msk_data_i.Text
    Mid(xLinha, 43, 10) = msk_data_f.Text
    BioImprime "@Printer.Print " & xLinha
    xLinha = "     Cartas-Frete adiante listadas, das quais relacionamos para fins de         "
    BioImprime "@Printer.Print " & xLinha
    xLinha = "     agendamento do pagamento a nosso favor, na geguinte conta bancária:        "
    BioImprime "@Printer.Print " & xLinha
    xLinha = "                    Banco..:                                                    "
    BioImprime "@Printer.Print " & xLinha
    xLinha = "                    Agência:                                                    "
    BioImprime "@Printer.Print " & xLinha
    xLinha = "                    Conta..:                                                    "
    BioImprime "@Printer.Print " & xLinha
    xLinha = "                    Titular:                                                    "
    Mid(xLinha, 30, 40) = g_nome_empresa
    BioImprime "@Printer.Print " & xLinha
    
    BioImprime "@Printer.Print " & "    "
    BioImprime "@Printer.Print " & "    "
    BioImprime "@Printer.Print " & "    "
    
    xLinha = "     RELAÇÃO DAS CARTAS-FRETE PAGAS:                                            "
    BioImprime "@Printer.Print " & xLinha
    
    
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@@Printer.FontBold = False"
    '''                                     10        20        30        40        50        60        70        80        90       100       110       120       130
    '''                             12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567
    
    
    BioImprime "@Printer.Print " & "     "
    
    BioImprime "@Printer.Print " & "         +------------+------------+------------------------------------------+--------------------------------+--------------+          "
    BioImprime "@Printer.Print " & "         |    DATA    |   NUMERO   | PROPRIETÁRIO                             | VEÍCULO                        |     VALOR    |"
    BioImprime "@Printer.Print " & "         |Troco em R$ |Troco Cheque| N.  Cheque |      Valor do Abastecimento |                                | CARTA  FRETE |"
    BioImprime "@Printer.Print " & "         +------------+------------+------------------------------------------+--------------------------------+--------------+          "
    
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i.Text = RetiraGString(1)
        msk_data_f.Text = RetiraGString(2)
        cmd_visualizar.SetFocus
    Else
        msk_data.Text = RetiraGString(1)
        msk_data_i.SetFocus
    End If
    g_string = " "
End Sub
Private Sub cmd_data_f_Click()
    g_string = msk_data_f.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i.Text = RetiraGString(1)
        msk_data_f.Text = RetiraGString(2)
    Else
        msk_data_f.Text = RetiraGString(1)
    End If
    g_string = " "
    txt_cliente.SetFocus
End Sub
Private Sub cmd_data_i_Click()
    g_string = msk_data_i.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i.Text = RetiraGString(1)
        msk_data_f.Text = RetiraGString(2)
        txt_cliente.SetFocus
    Else
        msk_data_i.Text = RetiraGString(1)
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
'    ElseIf Val(txt_cliente.Text) = 0 Then
'        MsgBox "Selecione um cliente.", vbInformation, "Atenção!"
'        txt_cliente.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_sair_Click()
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
        cmd_visualizar.SetFocus
    End If
End Sub
Private Sub dtcboCliente_LostFocus()
    If dtcboCliente.BoundText <> "" Then
        If Cliente.LocalizarCodigo(Val(dtcboCliente.BoundText)) Then
            If Cliente.Codigo <> Val(txt_cliente.Text) Then
                txt_cliente.Text = Cliente.Codigo
            End If
            cmd_visualizar.SetFocus
        End If
    End If
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    If Not IsDate(msk_data.Text) Then
        msk_data.Text = Format(g_data_def, "dd/mm/yyyy")
        msk_data_i.Text = Format(g_data_def - 1, "dd/mm/yyyy")
        msk_data_f.Text = Format(g_data_def - 1, "dd/mm/yyyy")
        msk_data_i.SetFocus
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
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_f_GotFocus()
    msk_data_f.SelStart = 0
    msk_data_f.SelLength = 5
End Sub
Private Sub msk_data_f_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_cliente.SetFocus
    End If
End Sub
Private Sub msk_data_i_GotFocus()
    msk_data_i.SelStart = 0
    msk_data_i.SelLength = 5
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
            If Cliente.Inativo = True Then
                MsgBox "O cliente " & Trim(Cliente.RazaoSocial) & " está inativo.", vbInformation, "Atenção!"
                dtcboCliente.BoundText = ""
                txt_cliente.SetFocus
                Exit Sub
            Else
                dtcboCliente.BoundText = Val(txt_cliente.Text)
                cmd_visualizar.SetFocus
                Exit Sub
            End If
        Else
            MsgBox "Cliente não cadastrado.", vbInformation, "Atenção!"
            txt_cliente.SetFocus
            Exit Sub
        End If
    Else
        dtcboCliente.BoundText = ""
    End If
End Sub
