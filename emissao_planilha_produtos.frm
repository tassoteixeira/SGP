VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form emissao_planilha_produtos 
   Caption         =   "Emissão da Planilha de Produtos"
   ClientHeight    =   3450
   ClientLeft      =   3990
   ClientTop       =   2010
   ClientWidth     =   6795
   Icon            =   "emissao_planilha_produtos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   Picture         =   "emissao_planilha_produtos.frx":030A
   ScaleHeight     =   3450
   ScaleWidth      =   6795
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1140
      Picture         =   "emissao_planilha_produtos.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Visualiza a Planilha de produtos."
      Top             =   2160
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3000
      Picture         =   "emissao_planilha_produtos.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Imprime a Planilha de produtos."
      Top             =   2160
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4860
      Picture         =   "emissao_planilha_produtos.frx":3074
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   2160
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   6555
      Begin VB.CheckBox chkImprimeBomba 
         Caption         =   "Imprime Bombas de Combustíveis"
         Height          =   255
         Left            =   1680
         TabIndex        =   10
         Top             =   1620
         Value           =   1  'Checked
         Width           =   3015
      End
      Begin VB.CheckBox chkExclusivoPosto 
         Caption         =   "Exclusivo do Posto"
         Height          =   255
         Left            =   1680
         TabIndex        =   7
         Top             =   1020
         Width           =   1875
      End
      Begin VB.CheckBox chkExclusivoLoja 
         Caption         =   "Exclusivo da Loja"
         Height          =   255
         Left            =   3660
         TabIndex        =   8
         Top             =   1020
         Width           =   1695
      End
      Begin VB.ComboBox cboSelecionar 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   600
         Width           =   4755
      End
      Begin VB.CheckBox chk_linha_separadora 
         Caption         =   "Imprime linha separadora"
         Height          =   255
         Left            =   1680
         TabIndex        =   9
         Top             =   1320
         Width           =   2235
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2760
         Picture         =   "emissao_planilha_produtos.frx":4706
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
      Begin VB.Label Label3 
         Caption         =   "I&mprimir Produto"
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   6
         Top             =   1020
         Width           =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "&Selecionar por"
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1515
      End
      Begin VB.Label Label5 
         Caption         =   "&Data de Emissão"
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1515
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   180
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "emissao_planilha_produtos"
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
Dim lNomeSubGrupo As String
Dim lCodigo As String
Dim lNome As String
Dim lUnidade As String
Dim lValor As String

Dim lSQL As String
Dim lRSCriado As Boolean
Private rsTabela As New adodb.Recordset
Private rsBomba As New adodb.Recordset
Dim rs As New adodb.Recordset
Dim rs2 As New adodb.Recordset

Private Produto As New cProduto
Private Sub CriaRS()
    With rs
        If lRSCriado Then
            .MoveFirst
            Do Until .EOF
                .Delete
                .MoveNext
            Loop
        Else
            .CursorLocation = adUseClient
            .Fields.Append "NomeSubGrupo", adVarChar, 40
            .Fields.Append "Codigo", adVarChar, 4
            .Fields.Append "Nome", adVarChar, 40
            .Fields.Append "Unidade", adVarChar, 3
            .Fields.Append "Valor", adVarChar, 10
            .Open
        End If
    End With
    With rs2
        If lRSCriado Then
            .MoveFirst
            Do Until .EOF
                .Delete
                .MoveNext
            Loop
        Else
            .CursorLocation = adUseClient
            .Fields.Append "Ordem", adVarChar, 4
            .Fields.Append "NomeSubGrupo", adVarChar, 40
            .Fields.Append "Codigo", adVarChar, 4
            .Fields.Append "Nome", adVarChar, 40
            .Fields.Append "Unidade", adVarChar, 3
            .Fields.Append "Valor", adVarChar, 10
            .Open
            lRSCriado = True
        End If
    End With
End Sub
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set Produto = Nothing
End Sub
Private Sub GravaRS()
    Dim i As Integer
    Dim i2 As Integer
    Dim xOrdem As String
    Do Until rsTabela.EOF
        If lNomeSubGrupo <> rsTabela("NomeSubGrupo").Value Then
            lNomeSubGrupo = rsTabela("NomeSubGrupo").Value
            rs.AddNew
            rs("NomeSubGrupo").Value = rsTabela("NomeSubGrupo").Value
            rs("Codigo").Value = 0
            rs("Nome").Value = ""
            rs("Unidade").Value = ""
            rs("Valor").Value = 0
            rs.Update
        End If
        rs.AddNew
        rs("NomeSubGrupo").Value = rsTabela("NomeSubGrupo").Value
        rs("Codigo").Value = rsTabela("Codigo").Value
        rs("Nome").Value = rsTabela("Nome").Value
        rs("Unidade").Value = rsTabela("Unidade").Value
        rs("Valor").Value = Format(rsTabela("Preco de Venda").Value, "0000000.00")
        rs.Update
        rsTabela.MoveNext
    Loop
    i = 0
    i2 = Format(rs.RecordCount / 2, "00000")
    xOrdem = "A"
    rs.MoveFirst
    Do Until rs.EOF
        i = i + 1
        If i > i2 Then
            xOrdem = "B"
            i = 1
        End If
        rs2.AddNew
        rs2("Ordem").Value = Format(i, "000") & xOrdem
        rs2("NomeSubGrupo").Value = rs("NomeSubGrupo").Value
        rs2("Codigo").Value = rs("Codigo").Value
        rs2("Nome").Value = rs("Nome").Value
        rs2("Unidade").Value = rs("Unidade").Value
        rs2("Valor").Value = Format(rs("Valor").Value, "0000000.00")
        rs2.Update
        rs.MoveNext
    Loop
End Sub
Private Sub ZeraVariaveis()
    lLinha = 0
    lPagina = 0
    lNomeSubGrupo = ""
End Sub
Private Sub PreencheCboSelecionar()
    cboSelecionar.Clear
    cboSelecionar.AddItem "Geral"
    cboSelecionar.ItemData(cboSelecionar.NewIndex) = 0
    cboSelecionar.AddItem "Menos Filtros"
    cboSelecionar.ItemData(cboSelecionar.NewIndex) = 1
    cboSelecionar.AddItem "Somente Filtros"
    cboSelecionar.ItemData(cboSelecionar.NewIndex) = 2
    cboSelecionar.AddItem "Nao Imprimir Produtos"
    cboSelecionar.ItemData(cboSelecionar.NewIndex) = 3
End Sub
Private Sub Relatorio()
    ZeraVariaveis
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "SELECT Produto.Nome, Produto.[Codigo do SubGrupo], Produto.Codigo, Estoque.[Preco de Venda], Produto.Unidade, Sub_Grupo.Nome as NomeSubGrupo"
    lSQL = lSQL & "  FROM Produto, Estoque, Sub_Grupo"
    lSQL = lSQL & " WHERE Estoque.Empresa = " & g_empresa
    lSQL = lSQL & "   AND Estoque.[Codigo do Produto2] = Produto.Codigo"
    lSQL = lSQL & "   AND Produto.Inativo = " & preparaBooleano(False)
    lSQL = lSQL & "   AND Produto.[Codigo do Grupo] <> 4"
    If cboSelecionar.ItemData(cboSelecionar.ListIndex) = 1 Then
        lSQL = lSQL & "   AND Produto.[Codigo do Grupo] <> 2"
    ElseIf cboSelecionar.ItemData(cboSelecionar.ListIndex) = 2 Then
        lSQL = lSQL & "   AND Produto.[Codigo do Grupo] = 2"
    End If
    lSQL = lSQL & "   AND Sub_Grupo.Codigo = Produto.[Codigo do SubGrupo]"
    If chkExclusivoPosto.Value = 1 And chkExclusivoLoja.Value = 0 Then
        lSQL = lSQL & "   AND Produto.[Exclusivo Posto] = " & preparaBooleano(True)
    End If
    If chkExclusivoLoja.Value = 1 And chkExclusivoPosto.Value = 0 Then
        lSQL = lSQL & "   AND Produto.[Exclusivo Loja] = " & preparaBooleano(True)
    End If
    lSQL = lSQL & " ORDER BY Sub_Grupo.Nome, Produto.Nome"
    'Abre RecordSet
    Set rsTabela = New adodb.Recordset
    Set rsTabela = Conectar.RsConexao(lSQL)
    'Verifica movimento
    If rsTabela.RecordCount > 0 Then
        CriaRS
        GravaRS
        ImpDados
    End If
    If rsTabela.State = 1 Then
        rsTabela.Close
    End If
    cmd_sair.SetFocus
End Sub
Private Sub ImpDados()
    LoopRS
    If lPagina > 0 Then
        ImpRodape
        If chkImprimeBomba.Value = 1 Then
            If cboSelecionar.ItemData(cboSelecionar.ListIndex) <> 3 Then
                BioImprime "@@Printer.NewPage"
                ImpCab
            End If
            ImpDetComposicaoCaixa
            ImpDetMedicao
        End If
        BioImprime "@@Printer.EndDoc"
        BioFechaImprime
        g_string = lLocal & lNomeArquivo & "|@|Relatório do Inventário de Produtos|@|"
        frm_preview.Show 1
    End If
End Sub
Private Sub LoopRS()
    'loop RS de Produtos
    Dim xLinha As String
    Dim i As Integer
    i = 0
    rs2.Sort = "Ordem, NomeSubGrupo, Nome"
    Do Until rs2.EOF
        If lPagina = 0 Then
            ImpCab
            If cboSelecionar.ItemData(cboSelecionar.ListIndex) <> 3 Then
                ImpCabProduto
            End If
        End If
        If lLinha >= 60 Then
            xLinha = "+----+------------------------------------+--------+-------+--------+----+------------------------------------+--------+-------+--------+"
            Mid(xLinha, 10, 22) = " Cerrado Informática. "
            BioImprime "@Printer.Print " & xLinha
            BioImprime "@@Printer.NewPage"
            ImpCab
            ImpCabProduto
        End If
        If i = 0 Then
            i = 1
            lNomeSubGrupo = rs2("NomeSubGrupo").Value
            lCodigo = rs2("Codigo").Value
            lNome = rs2("Nome").Value
            lUnidade = rs2("unidade").Value
            lValor = rs2("Valor").Value
        Else
            If cboSelecionar.ItemData(cboSelecionar.ListIndex) <> 3 Then
                ImpDet
            End If
            i = 0
        End If
        rs2.MoveNext
    Loop
    If i = 1 Then
        If cboSelecionar.ItemData(cboSelecionar.ListIndex) <> 3 Then
            ImpDet
        End If
    End If
End Sub
Private Sub ImpDet()
    Dim xLinha As String
    Dim i As Integer
    If chk_linha_separadora.Value = 1 Then
        xLinha = "+----+----------------------------+--------+------+-----+-----+-----+----+----------------------------+--------+------+-----+-----+-----+"
        BioImprime "@Printer.Print " & xLinha
        lLinha = lLinha + 1
    End If
    If chk_linha_separadora.Value = 1 Then
        xLinha = "|    |                            |        |      |     |     |     |    |                            |        |      |     |     |     |"
    Else
        xLinha = "|    |                            |        |______|_____|_____|_____|    |                            |        |______|_____|_____|_____|"
    End If
    If lNome = "" Then
        Mid(xLinha, 7, 28) = "** " & lNomeSubGrupo & " **"
    Else
        i = Len(Format(lCodigo, "#000"))
        Mid(xLinha, 2 + 4 - i, i) = Format(lCodigo, "#000")
        Mid(xLinha, 7, 28) = lNome
        i = Len(Format(lValor, "####0.00"))
        Mid(xLinha, 36 + 8 - i, i) = Format(lValor, "####0.00")
    End If
    If Not rs2.EOF Then
        If rs2("Nome").Value = "" Then
            Mid(xLinha, 75, 28) = "** " & rs2("NomeSubGrupo").Value & " **"
        Else
            i = Len(Format(rs2("Codigo").Value, "#000"))
            Mid(xLinha, 70 + 4 - i, i) = Format(rs2("Codigo").Value, "#000")
            Mid(xLinha, 75, 28) = rs2("Nome").Value
            i = Len(Format(rs2("Valor").Value, "####0.00"))
            Mid(xLinha, 104 + 8 - i, i) = Format(rs2("Valor").Value, "####0.00")
        End If
    End If
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 1
End Sub
Private Sub ImpDetBomba()
    Dim xLinha As String
    Dim i As Integer
    
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "SELECT Bomba.Codigo, Bomba.[Tipo de Combustivel], Bomba.[Preco de Venda], Bomba.[Tipo de Preco]"
    lSQL = lSQL & "  FROM Bomba"
    lSQL = lSQL & " WHERE Bomba.Empresa = " & g_empresa
    lSQL = lSQL & " ORDER BY Bomba.Codigo"
    'Abre RecordSet
    Set rsBomba = New adodb.Recordset
    Set rsBomba = Conectar.RsConexao(lSQL)
    If rsBomba.RecordCount > 0 Then
        BioImprime "@@Printer.FontName = Draft 10cpi"
        BioImprime "@Printer.Print " & "+--+-------------+-------------+---------+----------+--------------+--------+--+"
        BioImprime "@Printer.Print " & "|N.|   ABERTURA  |  ENCERRANTE |LTS.SAIDA|VLR LITRO |VALOR DA SAIDA| PREÇO  |CB|"
        BioImprime "@Printer.Print " & "+--+-------------+-------------+---------+----------+--------------+--------+--+"
        lLinha = lLinha + 3
        rsBomba.MoveFirst
        Do Until rsBomba.EOF
            xLinha = "|  |_____________|_____________|_________|__________|______________|        |  |"

            Mid(xLinha, 2, 2) = Format(rsBomba("Codigo").Value, "00")
            'i = Len(Format(x_abertura, "####,##0.0"))
            'Mid(x_linha, 5 + 10 - i, i) = Format(x_abertura, "####,##0.0")
            i = Len(Format(rsBomba("Preco de Venda").Value, "##,##0.000"))
            Mid(xLinha, 43 + 10 - i, i) = Format(rsBomba("Preco de Venda").Value, "##,##0.000")
            If rsBomba("Tipo de Preco").Value = "V" Then
                Mid(xLinha, 70, 7) = "A VISTA"
            Else
                Mid(xLinha, 70, 7) = "A PRAZO"
            End If
            Mid(xLinha, 78, 2) = rsBomba("Tipo de Combustivel").Value
            BioImprime "@Printer.Print " & xLinha
            lLinha = lLinha + 1
            rsBomba.MoveNext
        Loop
        BioImprime "@Printer.Print " & "+--+-------------+-------------+---------+----------+--------------+--------+--+"
        BioImprime "@Printer.Print " & "|                    TOTAL DA VENDA DE COMBUSTIVEIS |              |           |"
        BioImprime "@Printer.Print " & "+---------------------------------------------------+--------------+-----------+"
        lLinha = lLinha + 3
    End If
    rsBomba.Close
    Set rsBomba = Nothing
End Sub
Private Sub ImpDetComposicaoCaixa()
    Dim xLinha As String
    Dim i As Integer
    
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "SELECT Codigo, Nome"
    lSQL = lSQL & "  FROM Composicao_Caixa"
    lSQL = lSQL & " ORDER BY Ordem"
    'Abre RecordSet
    Set rsBomba = New adodb.Recordset
    Set rsBomba = Conectar.RsConexao(lSQL)
    If rsBomba.RecordCount > 0 Then
        BioImprime "@@Printer.FontName = Sans Serif 17cpi"
        BioImprime "@Printer.Print " & "+-----+--------------------------------+------------------+------------------+------------------+------------------+--------------------+"
        BioImprime "@Printer.Print " & "| COD | COMPOSIÇÃO DO CAIXA            |     1a SANGIA    |     2a SANGIA    |     3a SANGIA    |     4a SANGIA    |        TOTAL       |"
        BioImprime "@Printer.Print " & "+-----+--------------------------------+------------------+------------------+------------------+------------------+--------------------+"
        rsBomba.MoveFirst
        Do Until rsBomba.EOF
            xLinha = "|     |                                |                  |                  |                  |                  |                    |"
            Mid(xLinha, 3, 3) = Format(rsBomba("Codigo").Value, "000")
            Mid(xLinha, 9, 30) = rsBomba("Nome").Value
            BioImprime "@Printer.Print " & xLinha
            BioImprime "@Printer.Print " & "+-----+--------------------------------+------------------+------------------+------------------+------------------+--------------------+"
            rsBomba.MoveNext
        Loop
        BioImprime "@Printer.Print " & "|                                TOTAL |                  |                  |                  |                  |                    |"
        BioImprime "@Printer.Print " & "+-----+--------------------------------+------------------+------------------+------------------+------------------+--------------------+"
    End If
    rsBomba.Close
    Set rsBomba = Nothing
End Sub
Private Sub ImpDetMedicao()
    Dim xLinha As String
    BioImprime "@Printer.Print " & " "
    BioImprime "@Printer.Print " & "+----+----+-----------+-----------+----+----+-----------+-----------+----+----+-----------+-----------+----+----+-----------+-----------+"
    BioImprime "@Printer.Print " & "|TANQ|COMB| MED.REGUA |MED. LITROS|TANQ|COMB| MED.REGUA |MED. LITROS|TANQ|COMB| MED.REGUA |MED. LITROS|TANQ|COMB| MED.REGUA |MED. LITROS|"
    BioImprime "@Printer.Print " & "+----+----+-----------+-----------+----+----+-----------+-----------+----+----+-----------+-----------+----+----+-----------+-----------+"
    BioImprime "@Printer.Print " & "| 01 |    |           |           | 02 |    |           |           | 03 |    |           |           | 04 |    |           |           |"
    BioImprime "@Printer.Print " & "+----+----+-----------+-----------+----+----+-----------+-----------+----+----+-----------+-----------+----+----+-----------+-----------+"
    BioImprime "@Printer.Print " & "| 05 |    |           |           | 06 |    |           |           | 07 |    |           |           | 08 |    |           |           |"
    BioImprime "@Printer.Print " & "+----+----+-----------+-----------+----+----+-----------+-----------+----+----+-----------+-----------+----+----+-----------+-----------+"
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & " "
End Sub
Private Sub ImpRodape()
    Dim y_local As Single
    Dim xLinha As String
    Dim i As Integer
    
    If cboSelecionar.ItemData(cboSelecionar.ListIndex) <> 3 Then
        xLinha = "+----+----------------------------+--------+------+-----+-----+-----+----+----------------------------+--------+------+-----+-----+-----+"
        BioImprime "@Printer.Print " & xLinha
        xLinha = "|              TOTAL DOS PRODUTOS |                                      |                TOTAL GERAL |                                 |"
        BioImprime "@Printer.Print " & xLinha
        xLinha = "+---------------------------------+--------------------------------------+----------------------------+---------------------------------+"
        Mid(xLinha, 5, 22) = " Cerrado Informática. "
        BioImprime "@Printer.Print " & xLinha
    End If
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & " "
End Sub
Private Sub ImpCab()
    Dim xLinha As String
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
    BioImprime "@@Printer.Print " & Chr(34) & " " & Chr(34)
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@@Printer.CurrentY = 0"
    xLinha = "+---------------------------------------------------------------------------------------------------------------------------------------+"
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontBold = True"
    xLinha = "|                                                         PLANILHA DE PRODUTOS                                       CIDADE, __/__/____ |"
    Mid(xLinha, 3, 40) = g_nome_empresa
    i = Len(g_cidade_empresa)
    Mid(xLinha, 94 + 30 - i, i) = g_cidade_empresa
    Mid(xLinha, 126, 10) = msk_data.Text
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontBold = False"
    '                  1         2         3         4         5         6        x7         8         9        10        11        12        13     13
    '         12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567
    '                                                                                                            123456789012345678901234567890
    xLinha = "|                                                                                                                                       |"
    BioImprime "@Printer.Print " & xLinha
    
    xLinha = "| FRENTISTA:                                              DATA: _____/_____/_____                             PERÍODO: _____h ÀS _____h |"
    BioImprime "@Printer.Print " & xLinha
    If chkImprimeBomba.Value = 1 And lPagina = 1 Then
        ImpDetBomba
    End If
End Sub
Private Sub ImpCabProduto()
    Dim xLinha As String
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    xLinha = "+----+----------------------------+--------+------+-----+-----+-----+----+----------------------------+--------+------+-----+-----+-----+"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|COD.|DISCRIMINAÇÃO DOS PRODUTOS  |  VENDA |INICIO|ENTRA|SAIDA|SALDO|COD.|DISCRIMINAÇÃO DOS PRODUTOS  |  VENDA |INICIO|ENTRA|SAIDA|SALDO|"
    BioImprime "@Printer.Print " & xLinha
    If chk_linha_separadora.Value = 0 Then
        xLinha = "+----+----------------------------+--------+------+-----+-----+-----+----+----------------------------+--------+------+-----+-----+-----+"
        BioImprime "@Printer.Print " & xLinha
    End If
End Sub
Private Sub cboSelecionar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_visualizar.SetFocus
    End If
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    msk_data = RetiraGString(1)
    cboSelecionar.SetFocus
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
    ElseIf cboSelecionar.ListIndex = -1 Then
        MsgBox "Selecione o tipo de produto.", vbInformation, "Atenção!"
        cboSelecionar.SetFocus
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
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    If Not IsDate(msk_data.Text) Then
        msk_data.Text = Format(g_data_def, "dd/mm/yyyy")
        chkExclusivoPosto.Value = 1
        chkExclusivoLoja.Value = 0
        chkImprimeBomba.Value = 1
        cmd_visualizar.SetFocus
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
    lRSCriado = False
    PreencheCboSelecionar
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cboSelecionar.SetFocus
    End If
End Sub
