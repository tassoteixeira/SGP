VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form emissao_medida_tanque 
   Caption         =   "Emissão de Medida de Tanques"
   ClientHeight    =   2625
   ClientLeft      =   3990
   ClientTop       =   2010
   ClientWidth     =   6795
   Icon            =   "emissao_medida_tanque.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "emissao_medida_tanque.frx":030A
   ScaleHeight     =   2625
   ScaleWidth      =   6795
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1140
      Picture         =   "emissao_medida_tanque.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Visualiza a medida de tanque."
      Top             =   1680
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3000
      Picture         =   "emissao_medida_tanque.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Imprime a medida de tanque."
      Top             =   1680
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4860
      Picture         =   "emissao_medida_tanque.frx":3074
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   1680
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6555
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2760
         Picture         =   "emissao_medida_tanque.frx":4706
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   5940
         Picture         =   "emissao_medida_tanque.frx":59E0
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CheckBox chkVersaoEmail 
         Caption         =   "Versão para Email"
         Height          =   255
         Left            =   1680
         TabIndex        =   10
         Top             =   1080
         Width           =   3675
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2760
         Picture         =   "emissao_medida_tanque.frx":6CBA
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
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
      Begin VB.Label Label1 
         Caption         =   "&Data inicial"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1515
      End
      Begin VB.Label Label2 
         Caption         =   "D&ata final"
         Height          =   255
         Left            =   3840
         TabIndex        =   7
         Top             =   720
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
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "emissao_medida_tanque"
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
Dim lNomeEmpresa As String
Dim lSQL As String
Dim lSubTotalAnterior As Currency
Dim lSubTotalAtual As Currency
Dim lSubTotalEntrada As Currency
Dim lSubTotalVenda As Currency
Dim lSubTotalPerdaSobra As Currency
Dim lTotalAtual As Currency
Dim lQtdEmpresa As Integer
Dim lImprimiuNomeEmpresa As Boolean

Dim lTotMedidaAnterior As Currency
Dim lTotMedidaAtual As Currency
Dim lTotEntrada As Currency
Dim lTotVenda As Currency
Dim lTotPerdaSobra As Currency

Private EntradaCombustivel As New CadastroDLL.cEntradaCombustivel
Private MedicaoCombustivel As New CadastroDLL.cMedicaoCombustivel
Private MovimentoAfericao As New CadastroDLL.cMovimentoAfericao
Private MovimentoBomba As New CadastroDLL.cMovimentoBomba

Private rsMedidaTanque As New adodb.Recordset
Private Sub AgrupaArquivoEmail(ByVal pNomeArquivo As String)
    Dim xString As String

    Set gArquivoTXT = gArqTxt.OpenTextFile(pNomeArquivo, ForReading)
    Do Until gArquivoTXT.AtEndOfStream
        xString = gArquivoTXT.ReadLine
        BioImprime "@Printer.Print " & xString
    Loop
    gArquivoTXT.Close
    Set gArquivoTXT = Nothing
End Sub
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set EntradaCombustivel = Nothing
    Set MedicaoCombustivel = Nothing
    Set MovimentoAfericao = Nothing
    Set MovimentoBomba = Nothing
    Set rsMedidaTanque = Nothing
End Sub
Private Sub ZeraVariaveis()
    lLinha = 0
    lPagina = 0
    lSubTotalAnterior = 0
    lSubTotalAtual = 0
    lSubTotalEntrada = 0
    lSubTotalVenda = 0
    lSubTotalPerdaSobra = 0
    lTotalAtual = 0
    lQtdEmpresa = 0
    lImprimiuNomeEmpresa = False
    lTotMedidaAnterior = 0
    lTotMedidaAtual = 0
    lTotEntrada = 0
    lTotVenda = 0
    lTotPerdaSobra = 0
End Sub
Private Sub Relatorio()
    ZeraVariaveis
    
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "SELECT Empresas.Nome AS NomeEmpresa, Empresas.Codigo AS CodigoEmpresa, Combustivel.Nome, Combustivel.Codigo AS CodigoCombustivel, Sum(MedicaoCombustivel.Quantidade) As Quantidade"
    lSQL = lSQL & "  FROM MedicaoCombustivel, Empresas, Combustivel"
    lSQL = lSQL & " WHERE MedicaoCombustivel.Data = " & preparaData(CDate(msk_data_f.Text))
    lSQL = lSQL & "   AND Empresas.Codigo = MedicaoCombustivel.Empresa"
    lSQL = lSQL & "   AND Combustivel.Empresa = MedicaoCombustivel.Empresa"
    lSQL = lSQL & "   AND Combustivel.Codigo = MedicaoCombustivel.[Tipo de Combustivel]"
    'lSQL = lSQL & " ORDER BY MedicaoCombustivel.Empresa, MedicaoCombustivel.[Tipo de Combustivel]"
    lSQL = lSQL & " GROUP BY Empresas.Nome, Empresas.Codigo, Combustivel.Nome, Combustivel.Codigo"
    lSQL = lSQL & " ORDER BY Empresas.Nome, Combustivel.Nome"
    
    'Abre RecordSet
    Set rsMedidaTanque = New adodb.Recordset
    Set rsMedidaTanque = Conectar.RsConexao(lSQL)
    
    
    'Verifica movimento
    If rsMedidaTanque.RecordCount > 0 Then
        ImpDados
    End If
    If rsMedidaTanque.State = 1 Then
        rsMedidaTanque.Close
    End If
    If chkVersaoEmail.Value = False Then
        cmd_sair.SetFocus
    End If
End Sub
Private Sub ImpDados()
    Dim xLinha As String
    Dim xNomeArquivo As String
    Dim retval As Long
    'loop movimento de cheques
    Do Until rsMedidaTanque.EOF
        If lPagina = 0 Then
            ImpCab
            lNomeEmpresa = rsMedidaTanque("NomeEmpresa").Value
            'ImpCabEmpresa
        End If
        If rsMedidaTanque("NomeEmpresa").Value <> lNomeEmpresa Then
            If lSubTotalAnterior > 0 Or lSubTotalAtual > 0 Or lSubTotalEntrada > 0 Or lSubTotalVenda > 0 Then
                ImpSubTotal
            End If
            lNomeEmpresa = rsMedidaTanque("NomeEmpresa").Value
            lImprimiuNomeEmpresa = False
            'ImpCabEmpresa
        End If      '53
        If lLinha >= 62 Then
            If chkVersaoEmail.Value = False Then
                xLinha = "+------------------------------------------+----------------------+------------+------------+-----+-----+-------------------------------+"
                Mid(xLinha, 5, 22) = " Cerrado Informática. "
                BioImprime "@Printer.Print " & xLinha
                BioImprime "@@Printer.NewPage"
                ImpCab
            End If
        End If
        ImpDet
        rsMedidaTanque.MoveNext
    Loop
    'ImpSubTotal
    If lPagina > 0 Then
        If lSubTotalAtual > 0 Then
            ImpSubTotal
        End If
        ImpTotal
        BioImprime "@@Printer.EndDoc"
        BioFechaImprime
        g_string = lLocal & lNomeArquivo & "|@|Relatório de Medição de Combustíveis|@|"
        If chkVersaoEmail.Value Then
            If gArqTxt.FileExists("C:\EmailTemporario.html") Then
                Call gArqTxt.DeleteFile("C:\EmailTemporario.html", True)
            End If
            xNomeArquivo = lNomeArquivo
            Mid(xNomeArquivo, 12, 1) = "V"
            Call gArqTxt.CopyFile(xNomeArquivo, "C:\EmailTemporario.html")
            retval = Shell("C:\Arquivos de programas\Internet Explorer\IEXPLORE.EXE C:\EmailTemporario.html", vbNormalFocus)
        Else
            frm_preview.Show 1
        End If
    End If
End Sub
Private Sub ImpCabEmpresa()
    Dim xLinha As String
    Dim i As Integer
    '                  1         2         3         4         5         6         7         8         9        10        11        12        13     13
    '         12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567
    If chkVersaoEmail.Value = False Then
        xLinha = "|                                          |                      |            |            |            |            |            |    |"
        Mid(xLinha, 3, 40) = rsMedidaTanque("NomeEmpresa").Value
        BioImprime "@Printer.Print " & xLinha
        lLinha = lLinha + 1
    Else
        xLinha = " "
        BioImprime "@Printer.Print " & xLinha
        
        xLinha = "        <p><b><font size=" & Chr(34) & "4" & Chr(34) & ">"
        xLinha = xLinha & rsMedidaTanque("NomeEmpresa").Value
        xLinha = xLinha & "</font></b></p>"
        BioImprime "@Printer.Print " & xLinha
        
        xLinha = "        <p>"
        BioImprime "@Printer.Print " & xLinha
        
        xLinha = "          <font face=" & Chr(34) & "Courier" & Chr(34) & ">"
        BioImprime "@Printer.Print " & xLinha
    End If
End Sub
Private Sub ImpDet()
    Dim xLinha As String
    Dim i As Integer
    Dim xEstoqueAnterior As Currency
    Dim xEntrada As Currency
    Dim xSaida As Currency
    Dim xPerdaSobra As Currency
    '                  1         2         3         4         5         6         7         8         9        10        11        12        13     13
    '         12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567
        
    lSubTotalAtual = lSubTotalAtual + rsMedidaTanque("Quantidade").Value
    xEstoqueAnterior = MedicaoCombustivel.TotalMedidaCombustivel(rsMedidaTanque("CodigoEmpresa").Value, CDate(msk_data_i.Text), rsMedidaTanque("CodigoCombustivel").Value, 0)
    xEntrada = EntradaCombustivel.TotalEntradaPeriodo(rsMedidaTanque("CodigoEmpresa").Value, CDate(msk_data_i.Text), CDate(msk_data_f.Text) - 1, rsMedidaTanque("CodigoCombustivel").Value, 0)
    xSaida = MovimentoBomba.QuantidadeVendaData(rsMedidaTanque("CodigoEmpresa").Value, CDate(msk_data_i.Text), CDate(msk_data_f.Text) - 1, rsMedidaTanque("CodigoCombustivel").Value, 0)
    xPerdaSobra = rsMedidaTanque("Quantidade").Value - (xEstoqueAnterior + xEntrada - xSaida)
    lSubTotalAnterior = lSubTotalAnterior + xEstoqueAnterior
    lSubTotalEntrada = lSubTotalEntrada + xEntrada
    lSubTotalVenda = lSubTotalVenda + xSaida
    lSubTotalPerdaSobra = lSubTotalPerdaSobra + xPerdaSobra
    
    
    lTotMedidaAnterior = lTotMedidaAnterior + xEstoqueAnterior
    lTotMedidaAtual = lTotMedidaAtual + rsMedidaTanque("Quantidade").Value
    lTotEntrada = lTotEntrada + xEntrada
    lTotVenda = lTotVenda + xSaida
    lTotPerdaSobra = lTotPerdaSobra + xPerdaSobra
    
    
    If chkVersaoEmail.Value = False Then
        xLinha = "|                                          |                      |            |            |            |            |            |    |"
        If lImprimiuNomeEmpresa = False Then
            lImprimiuNomeEmpresa = True
            Mid(xLinha, 3, 40) = lNomeEmpresa
        End If
        Mid(xLinha, 49, 17) = rsMedidaTanque("Nome").Value
        i = Len(Format(xEstoqueAnterior, "######,##0"))
        Mid(xLinha, 69 + 10 - i, i) = Format(xEstoqueAnterior, "######,##0")
        i = Len(Format(xEntrada, "######,##0"))
        Mid(xLinha, 82 + 10 - i, i) = Format(xEntrada, "######,##0")
        i = Len(Format(xSaida, "######,##0"))
        Mid(xLinha, 95 + 10 - i, i) = Format(xSaida, "######,##0")
        i = Len(Format(rsMedidaTanque("Quantidade").Value, "######,##0"))
        Mid(xLinha, 108 + 10 - i, i) = Format(rsMedidaTanque("Quantidade").Value, "######,##0")
        i = Len(Format(xPerdaSobra, "######,##0;-#####,##0"))
        Mid(xLinha, 121 + 10 - i, i) = Format(xPerdaSobra, "######,##0;-#####,##0")
        BioImprime "@Printer.Print " & xLinha
        lLinha = lLinha + 1
    Else
        xLinha = "&nbsp;&nbsp;"
        xLinha = xLinha & rsMedidaTanque("Nome").Value
        For i = Len(Trim(rsMedidaTanque("Nome").Value)) + Len(Trim(Format(rsMedidaTanque("Quantidade").Value, "######,##0"))) To 33
            xLinha = xLinha & "&nbsp;"
        Next
        xLinha = xLinha & Format(xEstoqueAnterior, "#####,##0")
        For i = Len(Trim(Format(xEstoqueAnterior, "#####,##0"))) To 9
            xLinha = xLinha & "&nbsp;"
        Next
        xLinha = xLinha & Format(xEntrada, "#####,##0")
        For i = Len(Trim(Format(xEntrada, "#####,##0"))) To 9
            xLinha = xLinha & "&nbsp;"
        Next
        xLinha = xLinha & Format(xSaida, "#####,##0")
        For i = Len(Trim(Format(xSaida, "#####,##0"))) To 9
            xLinha = xLinha & "&nbsp;"
        Next
        xLinha = xLinha & Format(rsMedidaTanque("Quantidade").Value, "#####,##0")
        For i = Len(Trim(Format(rsMedidaTanque("Quantidade").Value, "#####,##0"))) To 9
            xLinha = xLinha & "&nbsp;"
        Next
        xLinha = xLinha & Format(xPerdaSobra, "#####,##0;-####,##0")
        xLinha = xLinha & "<br />"
        BioImprime "@Printer.Print " & xLinha
    End If
End Sub
Private Sub ImpSubTotal()
    Dim xLinha As String
    Dim i As Integer
    Dim xNome As String
        
    '                  1         2         3         4         5         6         7         8         9        10        11        12        13     13
    '         12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567
    
    lQtdEmpresa = lQtdEmpresa + 1
    If chkVersaoEmail.Value = False Then
        xLinha = "|                                          | * TOTAL COMBUSTIVEIS |            |            |            |            |            |    |"
        i = Len(Format(lSubTotalAnterior, "######,##0"))
        Mid(xLinha, 69 + 10 - i, i) = Format(lSubTotalAnterior, "######,##0")
        i = Len(Format(lSubTotalEntrada, "######,##0"))
        Mid(xLinha, 82 + 10 - i, i) = Format(lSubTotalEntrada, "######,##0")
        i = Len(Format(lSubTotalVenda, "######,##0"))
        Mid(xLinha, 95 + 10 - i, i) = Format(lSubTotalVenda, "######,##0")
        i = Len(Format(lSubTotalAtual, "######,##0"))
        Mid(xLinha, 108 + 10 - i, i) = Format(lSubTotalAtual, "######,##0")
        i = Len(Format(lSubTotalPerdaSobra, "######,##0;-#####,##0"))
        Mid(xLinha, 121 + 10 - i, i) = Format(lSubTotalPerdaSobra, "######,##0;-#####,##0")
        BioImprime "@Printer.Print " & xLinha
        xLinha = "+------------------------------------------+----------------------+------------+------------+------------+------------+------------+----+"
        BioImprime "@Printer.Print " & xLinha
        lLinha = lLinha + 1
    Else
        xNome = "Quantidade Total"
        xLinha = "&nbsp;&nbsp;"
        xLinha = xLinha & xNome
        For i = Len(Trim(xNome)) + Len(Trim(Format(lSubTotalAtual, "######,##0"))) To 33
            xLinha = xLinha & "&nbsp;"
        Next
        xLinha = xLinha & Format(lSubTotalAtual, "######,##0")
        xLinha = xLinha & "</font>"
        BioImprime "@Printer.Print " & xLinha
    
        xLinha = "        </p>"
        BioImprime "@Printer.Print " & xLinha
        
        xLinha = "        <hr />"
        BioImprime "@Printer.Print " & xLinha
    End If
    lSubTotalAnterior = 0
    lSubTotalAtual = 0
    lSubTotalEntrada = 0
    lSubTotalVenda = 0
    lSubTotalPerdaSobra = 0
End Sub
Private Sub ImpTotal()
    Dim xLinha As String
    Dim i As Integer
    
    If chkVersaoEmail.Value = False Then
        xLinha = "|                                          |   ** TOTAL GERAL **  |            |            |            |            |            |    |"
        i = Len(Format(lTotMedidaAnterior, "######,##0"))
        Mid(xLinha, 69 + 10 - i, i) = Format(lTotMedidaAnterior, "######,##0")
        i = Len(Format(lTotEntrada, "######,##0"))
        Mid(xLinha, 82 + 10 - i, i) = Format(lTotEntrada, "######,##0")
        i = Len(Format(lTotVenda, "######,##0"))
        Mid(xLinha, 95 + 10 - i, i) = Format(lTotVenda, "######,##0")
        i = Len(Format(lTotMedidaAtual, "######,##0"))
        Mid(xLinha, 108 + 10 - i, i) = Format(lTotMedidaAtual, "######,##0")
        i = Len(Format(lTotPerdaSobra, "######,##0;-#####,##0"))
        Mid(xLinha, 121 + 10 - i, i) = Format(lTotPerdaSobra, "######,##0;-#####,##0")
        BioImprime "@Printer.Print " & xLinha
        lLinha = lLinha + 1
        
        xLinha = "+------------------------------------------+----------------------+------------+------------+------------+------------+------------+----+"
        Mid(xLinha, 5, 22) = " Cerrado Informática. "
        Mid(xLinha, 55, 2) = Format(lQtdEmpresa, "00")
        BioImprime "@Printer.Print " & xLinha
        BioImprime "@@Printer.FontName = Draft 10cpi"
        BioImprime "@Printer.Print " & "  "
    Else
        AgrupaArquivoEmail ("C:\VB5\SGP\Data\" & "medicao combustivel rodape.htm")
    End If
End Sub
Private Sub ImpCab()
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
    If chkVersaoEmail.Value = False Then
        BioImprime "@@Printer.FontName = Draft 5cpi"
        BioImprime "@@Printer.FontName = Draft 10cpi"
        BioImprime "@@Printer.CurrentY = 0"
        BioImprime "@Printer.Print " & "+------------------------------------------------------------------------------+"
        xLinha = "|                                                                  Página, ___ |"
        Mid(xLinha, 3, 40) = g_nome_empresa
        Mid(xLinha, 76, 3) = Format(lPagina, "000")
        BioImprime "@Printer.Print " & xLinha
        xLinha = "| RELAÇÃO DAS MEDIDAS DE TANQUES                            CIDADE, __/__/____ |"
        i = Len(g_cidade_empresa)
        Mid(xLinha, 37 + 30 - i, i) = g_cidade_empresa
        Mid(xLinha, 69, 10) = msk_data.Text
        BioImprime "@Printer.Print " & xLinha
        xLinha = "| Referente a.: __/__/____ a __/__/____  (AS 06:00 HORAS)                      |"
        Mid(xLinha, 17, 10) = msk_data_i.Text
        Mid(xLinha, 30, 10) = msk_data_f.Text
        BioImprime "@Printer.Print " & xLinha
    End If
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@@Printer.FontBold = False"
    If chkVersaoEmail.Value = False Then
        BioImprime "@Printer.Print " & "+------------------------------------------+----------------------+------------+------------+------------+------------+------------+----+"
        BioImprime "@Printer.Print " & "| EMPRESA                                  | COMBUSTÍVEL          | QUANTIDADE | QUANTIDADE | QUANTIDADE | QUANTIDADE | PERDAS  OU |    |"
        xLinha = "|                                          |                      | __/__/____ | DE ENTRADA | DAS VENDAS | EM ESTOQUE | SOBRAS     |    |"
        Mid(xLinha, 69, 10) = msk_data_i.Text
        BioImprime "@Printer.Print " & xLinha
        BioImprime "@Printer.Print " & "+------------------------------------------+----------------------+------------+------------+------------+------------+------------+----+"
    Else
        AgrupaArquivoEmail ("C:\VB5\SGP\Data\" & "medicao combustivel cabecalho.htm")
        BioImprime "@Printer.Print " & "<h3 style=" & Chr(34) & "text-align:center" & Chr(34) & " >Referente a: " & msk_data_i.Text & " a: " & msk_data_f.Text & "</h3>"
    End If
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data.Text
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
    g_string = ""
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
    g_string = ""
    cmd_visualizar.SetFocus
End Sub
Private Sub cmd_data_i_Click()
    g_string = msk_data_i.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i.Text = RetiraGString(1)
        msk_data_f.Text = RetiraGString(2)
        cmd_visualizar.SetFocus
    Else
        msk_data_i.Text = RetiraGString(1)
        msk_data_f.SetFocus
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
        MsgBox "Data final deve ser maior ou igual a " & msk_data_i.Text & ".", vbInformation, "Atenção!"
        msk_data_f.SetFocus
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
        If chkVersaoEmail.Value = False Then
            If SelecionaImpressoraEpson(Me) Then
                Call GravaAuditoria(1, Me.name, 6, "")
                Relatorio
            End If
        Else
            cmd_sair.SetFocus
            Call GravaAuditoria(1, Me.name, 6, "")
            Relatorio
        End If
    End If
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    If Not IsDate(msk_data.Text) Then
        msk_data.Text = Format(g_data_def, "dd/mm/yyyy")
        msk_data_i.Text = Format(g_data_def - 1, "dd/mm/yyyy")
        msk_data_f.Text = Format(g_data_def, "dd/mm/yyyy")
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
    If g_nome_usuario = "L.M.C." Then
        Me.Caption = Me.Caption & " - LMC"
        EntradaCombustivel.NomeTabela = "Entrada_Combustivel_LMC"
        MedicaoCombustivel.NomeTabela = "MedicaoCombustivelLMC"
        MovimentoAfericao.NomeTabela = "Movimento_Afericao_LMC"
        MovimentoBomba.NomeTabela = "Movimento_Bomba_LMC"
    Else
        EntradaCombustivel.NomeTabela = "Entrada_Combustivel"
        MedicaoCombustivel.NomeTabela = "MedicaoCombustivel"
        MovimentoAfericao.NomeTabela = "Movimento_Afericao"
        MovimentoBomba.NomeTabela = "Movimento_Bomba"
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_f_GotFocus()
    msk_data_f.SelStart = 0
    msk_data_f.SelLength = 2
End Sub
Private Sub msk_data_f_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_visualizar.SetFocus
    End If
End Sub
Private Sub msk_data_GotFocus()
    msk_data.SelStart = 0
    msk_data.SelLength = 2
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
        cmd_visualizar.SetFocus
    End If
End Sub
