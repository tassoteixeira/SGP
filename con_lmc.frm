VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form consulta_lmc 
   Caption         =   "Consulta Dados do L.M.C."
   ClientHeight    =   6405
   ClientLeft      =   2910
   ClientTop       =   1740
   ClientWidth     =   6975
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "con_lmc.frx":0000
   ScaleHeight     =   6405
   ScaleWidth      =   6975
   Begin VB.TextBox txtZerarTanque 
      Height          =   315
      Left            =   1800
      MaxLength       =   1
      TabIndex        =   9
      Top             =   1620
      Width           =   435
   End
   Begin VB.CommandButton cmd_mais 
      Caption         =   "&+"
      Height          =   855
      Left            =   3600
      Picture         =   "con_lmc.frx":0046
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Visualiza resumo da data posterior."
      Top             =   2220
      Width           =   735
   End
   Begin VB.CommandButton cmd_menos 
      Caption         =   "&-"
      Height          =   855
      Left            =   2640
      Picture         =   "con_lmc.frx":1438
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Visualiza resumo da data anterior."
      Top             =   2220
      Width           =   735
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   1680
      Picture         =   "con_lmc.frx":282A
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Visualiza o resumo do lmc."
      Top             =   2220
      Width           =   735
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4560
      Picture         =   "con_lmc.frx":3E34
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   2220
      Width           =   735
   End
   Begin VB.Frame frm_dados 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2700
         Picture         =   "con_lmc.frx":54C6
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.ComboBox cbo_distribuicao 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1080
         Width           =   2115
      End
      Begin VB.ComboBox cbo_combustivel 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   660
         Width           =   4935
      End
      Begin MSMask.MaskEdBox msk_data_i 
         Height          =   315
         Left            =   1680
         TabIndex        =   2
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label10 
         Caption         =   "&Tanque a Zerar"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label lbl_distribuicao 
         Caption         =   "&Fazer distribuição"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1140
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "&Data inicial"
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1515
      End
      Begin VB.Label Label2 
         Caption         =   "&Combustível"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   660
         Width           =   1455
      End
   End
   Begin VB.Frame frmDados 
      Caption         =   "Resumo do Período Informado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2955
      Left            =   1320
      TabIndex        =   14
      Top             =   3240
      Width           =   4395
      Begin VB.TextBox txt_perdas_sobras 
         Height          =   315
         Left            =   2520
         TabIndex        =   28
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label lbl_afericao 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2520
         TabIndex        =   22
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Total das Aferições"
         Height          =   255
         Left            =   180
         TabIndex        =   21
         Top             =   1500
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "Estoque de Abertura"
         Height          =   255
         Left            =   180
         TabIndex        =   15
         Top             =   420
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "Total Recebido"
         Height          =   255
         Left            =   180
         TabIndex        =   17
         Top             =   780
         Width           =   2295
      End
      Begin VB.Label Label6 
         Caption         =   "Total das Vendas"
         Height          =   255
         Left            =   180
         TabIndex        =   19
         Top             =   1140
         Width           =   2295
      End
      Begin VB.Label Label7 
         Caption         =   "Estoque Escritural"
         Height          =   255
         Left            =   180
         TabIndex        =   23
         Top             =   1860
         Width           =   2295
      End
      Begin VB.Label Label8 
         Caption         =   "Estoque Fechamento"
         Height          =   255
         Left            =   180
         TabIndex        =   25
         Top             =   2220
         Width           =   2295
      End
      Begin VB.Label Label9 
         Caption         =   "- &Perdas + Sobras"
         Height          =   255
         Left            =   180
         TabIndex        =   27
         Top             =   2580
         Width           =   2295
      End
      Begin VB.Label lbl_estoque_abertura 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2520
         TabIndex        =   16
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lbl_total_recebido 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2520
         TabIndex        =   18
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label lbl_vendas_dia 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2520
         TabIndex        =   20
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lbl_estoque_escritural 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2520
         TabIndex        =   24
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label lbl_estoque_fechamento 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2520
         TabIndex        =   26
         Top             =   2160
         Width           =   1695
      End
   End
End
Attribute VB_Name = "consulta_lmc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_tipo_combustivel As String
Dim l_data_i As Date
Dim l_abertura_tanque As Currency
Dim l_fechamento_tanque As Currency
Dim l_quantidade_entrada As Currency
Dim lQuantidadeAfericao As Currency
Dim l_litros_vendidos As Currency
Dim l_existe As Boolean
Dim l_litros_perdas_sobras As Currency
Dim l_operacao As String
Dim l_processamento As Boolean
Dim lSQL As String
Dim lQuantidadeTanque As Currency
Dim lCodigoTanque(0 To 10) As Integer
Dim lEstoqueTanque(0 To 10) As Currency
Dim lLimiteAcumulPerdasSobras As Integer
Dim lLimiteAcumuladoPerdas As Currency
Dim lLimiteAcumuladoSobras As Currency
Dim lLimiteDiaPerdasSobras As Integer

Private Combustivel As New cCombustivel
Private ConfiguracaoDiversa As New cConfiguracaoDiversa
Private EntradaCombustivel As New cEntradaCombustivel
Private LivroLMC As New cLivroLMC
Private MedicaoCombustivel As New cMedicaoCombustivel
Private MovimentoBomba As New cMovimentoBomba
Private TanqueCombustivel As New cTanqueCombustivel

Dim rstAfericao As New adodb.Recordset
Dim rstTanques As New adodb.Recordset
Dim rsEntradaCombustivel As New adodb.Recordset
Dim rsMovimentoBomba As New adodb.Recordset
Private Sub ZeraVariaveis()
    l_abertura_tanque = 0
    l_fechamento_tanque = 0
    l_quantidade_entrada = 0
    lQuantidadeAfericao = 0
    l_litros_vendidos = 0
End Sub
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set Combustivel = Nothing
    Set ConfiguracaoDiversa = Nothing
    Set EntradaCombustivel = Nothing
    Set LivroLMC = Nothing
    Set MedicaoCombustivel = Nothing
    Set MovimentoBomba = Nothing
    Set TanqueCombustivel = Nothing
End Sub
Private Sub CalculaEstoque()
    Dim xPorcentagem As Integer
    Dim xPorcentagemAcumulada As Integer
    Dim xMaior As Integer
    Dim xMenor As Integer
    
    If lQuantidadeTanque = 2 Then
        xMaior = 90
        xMenor = 10
        xPorcentagem = Int((xMaior - xMenor + 1) * Rnd + xMenor)
        lEstoqueTanque(1) = Format(lbl_estoque_fechamento.Caption * xPorcentagem / 100, "0000000000")
        xPorcentagemAcumulada = xPorcentagem
        xPorcentagem = 100 - xPorcentagemAcumulada
        lEstoqueTanque(2) = Format(lbl_estoque_fechamento.Caption * xPorcentagem / 100, "0000000000")
        xPorcentagemAcumulada = xPorcentagemAcumulada + xPorcentagem
    ElseIf lQuantidadeTanque = 3 Then
        xMaior = 70
        xMenor = 10
        xPorcentagem = Int((xMaior - xMenor + 1) * Rnd + xMenor)
        lEstoqueTanque(1) = Format(lbl_estoque_fechamento.Caption * xPorcentagem / 100, "0000000000")
        xPorcentagemAcumulada = xPorcentagem
        
        xMaior = 100 - xPorcentagemAcumulada - 10
        xPorcentagem = Int((xMaior - xMenor + 1) * Rnd + xMenor)
        lEstoqueTanque(2) = Format(lbl_estoque_fechamento.Caption * xPorcentagem / 100, "0000000000")
        xPorcentagemAcumulada = xPorcentagemAcumulada + xPorcentagem
        
        xPorcentagem = 100 - xPorcentagemAcumulada
        lEstoqueTanque(3) = Format(lbl_estoque_fechamento.Caption * xPorcentagem / 100, "0000000000")
        xPorcentagemAcumulada = xPorcentagemAcumulada + xPorcentagem
    ElseIf lQuantidadeTanque = 4 Then
        xMaior = 60
        xMenor = 10
        xPorcentagem = Int((xMaior - xMenor + 1) * Rnd + xMenor)
        lEstoqueTanque(1) = Format(lbl_estoque_fechamento.Caption * xPorcentagem / 100, "0000000000")
        xPorcentagemAcumulada = xPorcentagem
        
        xMaior = 100 - xPorcentagemAcumulada - 20
        xPorcentagem = Int((xMaior - xMenor + 1) * Rnd + xMenor)
        lEstoqueTanque(2) = Format(lbl_estoque_fechamento.Caption * xPorcentagem / 100, "0000000000")
        xPorcentagemAcumulada = xPorcentagemAcumulada + xPorcentagem
        
        xMaior = 100 - xPorcentagemAcumulada - 10
        xPorcentagem = Int((xMaior - xMenor + 1) * Rnd + xMenor)
        lEstoqueTanque(3) = Format(lbl_estoque_fechamento.Caption * xPorcentagem / 100, "0000000000")
        xPorcentagemAcumulada = xPorcentagemAcumulada + xPorcentagem
        
        xPorcentagem = 100 - xPorcentagemAcumulada
        lEstoqueTanque(4) = Format(lbl_estoque_fechamento.Caption * xPorcentagem / 100, "0000000000")
        xPorcentagemAcumulada = xPorcentagemAcumulada + xPorcentagem
    End If
End Sub
Function CalculaPerdasSobras() As Currency
    Dim xUnidade As Integer
    Dim xDezena As Integer
    Dim xCentena As Integer
    
    'Aqui sorteia se será negativo ou positivo
    CalculaPerdasSobras = Int((2 * Rnd) + 1)
    If CalculaPerdasSobras = 1 Then
        l_operacao = "+"
    Else
        l_operacao = "-"
    End If
    
    'Se o acumulado das sobras for maior que 199
    'Então passa a ser negativo
    'If l_litros_perdas_sobras > 199 Then
    If l_litros_perdas_sobras > lLimiteAcumulPerdasSobras Then
        l_operacao = "-"
    End If
    
    'Se o acumulado das perdas for menor que -199
    'Então passa a ser positivo
    'If l_litros_perdas_sobras < -199 Then
    If l_litros_perdas_sobras < -lLimiteAcumulPerdasSobras Then
        l_operacao = "+"
    End If
    
    If cbo_distribuicao = "+" Then
        l_operacao = "+"
    ElseIf cbo_distribuicao = "-" Then
        l_operacao = "-"
    End If
    
    
    If lLimiteAcumulPerdasSobras = 299 Then
        xUnidade = Int((9 * Rnd) + 1)
        xDezena = Int((9 * Rnd) + 1)
        xCentena = Int((2 * Rnd) + 0)
    ElseIf lLimiteAcumulPerdasSobras = 249 Then
        xUnidade = Int((9 * Rnd) + 1)
        xCentena = Int((2 * Rnd) + 0)
        If xCentena = 1 Then
            xDezena = Int((6 * Rnd) + 1)
        Else
            xDezena = Int((9 * Rnd) + 1)
        End If
    ElseIf lLimiteAcumulPerdasSobras = 199 Then
        xUnidade = Int((9 * Rnd) + 1)
        xCentena = Int((2 * Rnd) + 0)
        If xCentena = 1 Then
            xDezena = Int((3 * Rnd) + 1)
        Else
            xDezena = Int((9 * Rnd) + 1)
        End If
    ElseIf lLimiteAcumulPerdasSobras = 149 Then
        xUnidade = Int((9 * Rnd) + 1)
        xDezena = Int((9 * Rnd) + 1)
        xCentena = 0
    ElseIf lLimiteAcumulPerdasSobras = 99 Then
        xUnidade = Int((9 * Rnd) + 1)
        xDezena = Int((4 * Rnd) + 1)
        xCentena = 0
    ElseIf lLimiteAcumulPerdasSobras = 49 Then
        xUnidade = Int((9 * Rnd) + 1)
        xDezena = Int((2 * Rnd) + 1)
        xCentena = 0
    ElseIf lLimiteAcumulPerdasSobras = 19 Then
        xUnidade = Int((9 * Rnd) + 1)
        xDezena = 0
        xCentena = 0
    End If
    CalculaPerdasSobras = Val(CStr(xCentena) & CStr(xDezena) & CStr(xUnidade))
    If l_operacao = "+" Then
        l_litros_perdas_sobras = l_litros_perdas_sobras + CalculaPerdasSobras
    Else
        l_litros_perdas_sobras = l_litros_perdas_sobras - CalculaPerdasSobras
    End If
    If l_operacao = "-" Then
        CalculaPerdasSobras = (CalculaPerdasSobras - CalculaPerdasSobras * 2) - CCur("0," & Mid(lbl_estoque_escritural, Len(lbl_estoque_escritural) - 1, 2))
    Else
        CalculaPerdasSobras = CalculaPerdasSobras + CCur("0," & Mid(lbl_vendas_dia, Len(lbl_vendas_dia) - 1, 2))
    End If
End Function
Private Sub Consulta()
    ZeraVariaveis
    'Lê medição de combustível de Abertura
    If MedicaoCombustivel.LocalizarPrimeiroTanqueComb(g_empresa, l_data_i, l_tipo_combustivel) Then
        l_abertura_tanque = MedicaoCombustivel.Quantidade
        Do Until MedicaoCombustivel.LocalizarProximoTanqueComb(g_empresa, l_data_i, l_tipo_combustivel) = False
            l_abertura_tanque = l_abertura_tanque + MedicaoCombustivel.Quantidade
        Loop
    Else
        MsgBox "Não existe medição de combustíveis de abertura nesta data!", vbInformation, "Atenção!"
    End If
    
    
    
    'Lê medição de combustível de Fechamento
    If MedicaoCombustivel.LocalizarPrimeiroTanqueComb(g_empresa, l_data_i + 1, l_tipo_combustivel) Then
        l_fechamento_tanque = MedicaoCombustivel.Quantidade
        l_existe = True
        Do Until MedicaoCombustivel.LocalizarProximoTanqueComb(g_empresa, l_data_i + 1, l_tipo_combustivel) = False
            l_fechamento_tanque = l_fechamento_tanque + MedicaoCombustivel.Quantidade
        Loop
    Else
        l_existe = False
        MsgBox "Não existe medição de combustíveis de fechamento para esta data!", vbInformation, "Atenção!"
    End If
    
    
    'lê arquivo de combustível
    If Not Combustivel.LocalizarCodigo(g_empresa, l_tipo_combustivel) Then
        MsgBox "Combustível inexistente", vbInformation, "Erro de Consistência de Dados"
        cbo_combustivel.SetFocus
        Exit Sub
    End If
    
    'lê entradas de combustíveis
    
    'cbo_combustivel.Clear
    'Prepara SQL
    lSQL = "SELECT SUM(Quantidade) AS TotalQantidade"
    lSQL = lSQL & "  FROM " & EntradaCombustivel.NomeTabela
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    lSQL = lSQL & "   AND [Tipo de Combustivel] = " & preparaTexto(l_tipo_combustivel)
    lSQL = lSQL & "   AND Data = " & preparaData(l_data_i)
    'Abre RecordSet
    Set rsEntradaCombustivel = Conectar.RsConexao(lSQL)
    'Verifica tabela
    If rsEntradaCombustivel.RecordCount > 0 Then
        If Not IsNull(rsEntradaCombustivel("TotalQantidade").Value) Then
            l_quantidade_entrada = rsEntradaCombustivel("TotalQantidade").Value
        End If
    End If
    rsEntradaCombustivel.Close
    Set rsEntradaCombustivel = Nothing
    
    
    
    'Lê movimentação das Aferições
    lSQL = "SELECT SUM(Quantidade) AS Total FROM Movimento_Afericao_LMC"
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    lSQL = lSQL & " AND [Tipo de Combustivel] = " & preparaTexto(l_tipo_combustivel)
    lSQL = lSQL & " AND Data = " & preparaData(l_data_i)
    Set rstAfericao = Conectar.RsConexao(lSQL)
    If Not rstAfericao.EOF Then
        If Not IsNull(rstAfericao!total) Then
            lQuantidadeAfericao = rstAfericao!total
        End If
    End If
    rstAfericao.Close
    Set rstAfericao = Nothing

    
    
    'lê movimentação das bombas
    lSQL = "SELECT SUM([Quantidade da Saida]) AS Total"
    lSQL = lSQL & " FROM Movimento_Bomba_LMC"
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    lSQL = lSQL & " AND [Tipo de Combustivel] = " & preparaTexto(l_tipo_combustivel)
    lSQL = lSQL & " AND Data = " & preparaData(l_data_i)
    Set rsMovimentoBomba = Conectar.RsConexao(lSQL)
    If Not rsMovimentoBomba.EOF Then
        If Not IsNull(rsMovimentoBomba!total) Then
            l_litros_vendidos = rsMovimentoBomba!total
        End If
    End If
    rsMovimentoBomba.Close
    Set rsMovimentoBomba = Nothing
'    tbl_movimento_bomba.Index = "id_data_tipo_combustivel"
'    tbl_movimento_bomba.Seek ">", g_empresa, l_data_i, Trim(l_tipo_combustivel), "0", "0"
'    If Not tbl_movimento_bomba.NoMatch Then
'        Do Until tbl_movimento_bomba.EOF
'            If tbl_movimento_bomba!Empresa <> g_empresa Then
'                Exit Do
'            End If
'            If tbl_movimento_bomba!Data <> l_data_I Then
'                Exit Do
'            End If
'            If Trim(tbl_movimento_bomba![Tipo de Combustivel]) = Trim(l_tipo_combustivel) Then
'                l_litros_vendidos = l_litros_vendidos + tbl_movimento_bomba![Quantidade da Saida]
'            End If
'            tbl_movimento_bomba.MoveNext
'        Loop
'    End If
    lbl_estoque_abertura.Caption = Format(l_abertura_tanque, "###,##0.00")
    lbl_total_recebido.Caption = Format(l_quantidade_entrada, "###,##0.00")
    lbl_vendas_dia.Caption = Format(l_litros_vendidos, "###,##0.00")
    lbl_afericao.Caption = Format(lQuantidadeAfericao, "###,##0.00")
    lbl_estoque_escritural.Caption = Format((l_abertura_tanque + l_quantidade_entrada + lQuantidadeAfericao - l_litros_vendidos), "###,##0.00")
    lbl_estoque_fechamento.Caption = Format(l_fechamento_tanque, "###,##0.00")
    txt_perdas_sobras.Text = Format((l_fechamento_tanque - (l_abertura_tanque + l_quantidade_entrada + lQuantidadeAfericao - l_litros_vendidos)), "###,##0.00;-###,##0.00")
    If Mid(lbl_estoque_fechamento.Caption, Len(lbl_estoque_fechamento.Caption), 1) <> 0 Then
        MsgBox "Estoque de fechamento tem que ser numero redondo."
    End If
    If Val(lbl_estoque_fechamento.Caption) < 0 Then
        lbl_estoque_fechamento.BackColor = vbRed
        MsgBox "Estoque de fechamento tem que ser positivo."
    Else
        lbl_estoque_fechamento.BackColor = 16777215
    End If
    If Not l_existe Then
        lbl_estoque_fechamento.Caption = Format((l_abertura_tanque + l_quantidade_entrada + lQuantidadeAfericao - l_litros_vendidos), "###,##0.00")
        txt_perdas_sobras.Text = Format(0, "###,##0.00;-###,##0.00")
        GravaMedicao
    Else
        txt_perdas_sobras.SelLength = Len(txt_perdas_sobras.Text)
        txt_perdas_sobras.SetFocus
    End If
End Sub
Private Sub Altera()
    Dim i As Integer
    Dim i2 As Integer
    Dim xStringLog As String
    
    On Error GoTo FileError
    
    If l_processamento Then
        txt_perdas_sobras.Text = CalculaPerdasSobras
    End If
    txt_perdas_sobras.Text = Format(txt_perdas_sobras.Text, "###,##0.00;-###,##0.00")
    lbl_estoque_fechamento.Caption = Format(l_abertura_tanque + l_quantidade_entrada + lQuantidadeAfericao - l_litros_vendidos + fValidaValor2(txt_perdas_sobras.Text), "###,###.00")
    
    
    If MedicaoCombustivel.LocalizarPrimeiroTanqueComb(g_empresa, l_data_i + 1, l_tipo_combustivel) Then
        If lQuantidadeTanque > 1 Then
            CalculaEstoque
            For i = 1 To lQuantidadeTanque
                If MedicaoCombustivel.LocalizarCodigo(g_empresa, l_data_i + 1, lCodigoTanque(i)) Then
                    If lCodigoTanque(i) = Val(txtZerarTanque.Text) Then
                        xStringLog = "De Data:" & Format(MedicaoCombustivel.Data, "dd/mm/yyyy") & " Comb:" & MedicaoCombustivel.TipoCombustivel & " Tq:" & MedicaoCombustivel.NumeroTanque & " Med:" & MedicaoCombustivel.Quantidade
                        Call GravaAuditoria(1, Me.name, 26, xStringLog)
                        MedicaoCombustivel.Quantidade = 0
                        xStringLog = "Para Data:" & Format(MedicaoCombustivel.Data, "dd/mm/yyyy") & " Comb:" & MedicaoCombustivel.TipoCombustivel & " Tq:" & MedicaoCombustivel.NumeroTanque & " Med:" & MedicaoCombustivel.Quantidade
                        Call GravaAuditoria(1, Me.name, 26, xStringLog)
                        If Not MedicaoCombustivel.Alterar(g_empresa, l_data_i + 1, lCodigoTanque(i)) Then
                            MsgBox "Não foi possível alterar a medição de tanque.", vbInformation, "Erro de Integridade"
                        End If
                    Else
                        If Val(txtZerarTanque.Text) > 0 Then
                            For i2 = 1 To lQuantidadeTanque
                                If lCodigoTanque(i2) <> lCodigoTanque(i) Then
                                    lEstoqueTanque(i) = lEstoqueTanque(i) + lEstoqueTanque(i2)
                                End If
                            Next
                        End If
                        xStringLog = "De Data:" & Format(MedicaoCombustivel.Data, "dd/mm/yyyy") & " Comb:" & MedicaoCombustivel.TipoCombustivel & " Tq:" & MedicaoCombustivel.NumeroTanque & " Med:" & MedicaoCombustivel.Quantidade
                        Call GravaAuditoria(1, Me.name, 26, xStringLog)
                        MedicaoCombustivel.Quantidade = lEstoqueTanque(i)
                        xStringLog = "Para Data:" & Format(MedicaoCombustivel.Data, "dd/mm/yyyy") & " Comb:" & MedicaoCombustivel.TipoCombustivel & " Tq:" & MedicaoCombustivel.NumeroTanque & " Med:" & MedicaoCombustivel.Quantidade
                        Call GravaAuditoria(1, Me.name, 26, xStringLog)
                        If Not MedicaoCombustivel.Alterar(g_empresa, l_data_i + 1, lCodigoTanque(i)) Then
                            MsgBox "Não foi possível alterar a medição de tanque.", vbInformation, "Erro de Integridade"
                        End If
                    End If
                End If
            Next
        Else
            xStringLog = "De Data:" & Format(MedicaoCombustivel.Data, "dd/mm/yyyy") & " Comb:" & MedicaoCombustivel.TipoCombustivel & " Tq:" & MedicaoCombustivel.NumeroTanque & " Med:" & MedicaoCombustivel.Quantidade
            Call GravaAuditoria(1, Me.name, 26, xStringLog)
            MedicaoCombustivel.Quantidade = fValidaValor2(lbl_estoque_fechamento.Caption)
            xStringLog = "Para Data:" & Format(MedicaoCombustivel.Data, "dd/mm/yyyy") & " Comb:" & MedicaoCombustivel.TipoCombustivel & " Tq:" & MedicaoCombustivel.NumeroTanque & " Med:" & MedicaoCombustivel.Quantidade
            Call GravaAuditoria(1, Me.name, 26, xStringLog)
            If Not MedicaoCombustivel.Alterar(g_empresa, l_data_i + 1, MedicaoCombustivel.NumeroTanque) Then
                MsgBox "Não foi possível alterar a medição de tanque.", vbInformation, "Erro de Integridade"
            End If
        End If
    Else
        If TanqueCombustivel.LocalizarPrimeiroTanqueComb(g_empresa, l_tipo_combustivel) Then
            MedicaoCombustivel.Empresa = g_empresa
            MedicaoCombustivel.Data = Format(CDate(l_data_i + 1), "dd/mm/yyyy")
            MedicaoCombustivel.NumeroTanque = TanqueCombustivel.NumeroTanque
            MedicaoCombustivel.TipoCombustivel = l_tipo_combustivel
            MedicaoCombustivel.Quantidade = fValidaValor2(lbl_estoque_fechamento.Caption)
            MedicaoCombustivel.Observacao1 = ""
            MedicaoCombustivel.Observacao2 = ""
            MedicaoCombustivel.Observacao3 = ""
            MedicaoCombustivel.DescontoDiaAnterior = 0
            xStringLog = "Criada Data:" & Format(MedicaoCombustivel.Data, "dd/mm/yyyy") & " Comb:" & MedicaoCombustivel.TipoCombustivel & " Tq:" & MedicaoCombustivel.NumeroTanque & " Med:" & MedicaoCombustivel.Quantidade
            Call GravaAuditoria(1, Me.name, 26, xStringLog)
            If Not MedicaoCombustivel.Incluir Then
                MsgBox "Não foi possível incluir a medição de tanque.", vbInformation, "Erro de Integridade"
            End If
        Else
            MsgBox "Não foi possível encontrar um tanque para este combustível." & Chr(10) & "Não foi possível incluir a medição de tanque.", vbInformation, "Erro de Integridade"
        End If
    End If
    Exit Sub
FileError:
    'ErroArquivo tbl_medicao_combustivel.Name, "Medição de Combustívela"
    Exit Sub
End Sub
Private Sub GravaMedicao()
    On Error GoTo FileError
    If TanqueCombustivel.LocalizarPrimeiroTanqueComb(g_empresa, l_tipo_combustivel) Then
        MedicaoCombustivel.Empresa = g_empresa
        MedicaoCombustivel.Data = Format(CDate(l_data_i + 1), "dd/mm/yyyy")
        MedicaoCombustivel.NumeroTanque = TanqueCombustivel.NumeroTanque
        MedicaoCombustivel.TipoCombustivel = l_tipo_combustivel
        MedicaoCombustivel.Quantidade = fValidaValor2(lbl_estoque_fechamento.Caption)
        MedicaoCombustivel.Observacao1 = ""
        MedicaoCombustivel.Observacao2 = ""
        MedicaoCombustivel.Observacao3 = ""
        MedicaoCombustivel.DescontoDiaAnterior = 0
        If Not MedicaoCombustivel.Incluir Then
            MsgBox "Não foi possível incluir a medição de tanque.", vbInformation, "Erro de Integridade"
        End If
    Else
        MsgBox "Não foi possível encontrar um tanque para este combustível." & Chr(10) & "Não foi possível incluir a medição de tanque.", vbInformation, "Erro de Integridade"
    End If
    'tbl_medicao_combustivel.AddNew
    'tbl_medicao_combustivel!Empresa = g_empresa
    'tbl_medicao_combustivel!Data = l_data_I + 1
    'tbl_medicao_combustivel!tipo_combustivel = l_tipo_combustivel
    'tbl_medicao_combustivel!tanque_1 = fValidaValor2(lbl_estoque_fechamento.Caption)
    'tbl_medicao_combustivel!tanque_2 = 0
    'tbl_medicao_combustivel!tanque_3 = 0
    'tbl_medicao_combustivel!tanque_4 = 0
    'tbl_medicao_combustivel!tanque_5 = 0
    'tbl_medicao_combustivel!tanque_6 = 0
    'tbl_medicao_combustivel!observacao_1 = ""
    'tbl_medicao_combustivel!observacao_2 = ""
    'tbl_medicao_combustivel!observacao_3 = ""
    'tbl_medicao_combustivel![Desconto Dia Anterior] = 0
    'tbl_medicao_combustivel.Update
    Exit Sub
FileError:
    'ErroArquivo tbl_medicao_combustivel.Name, "Medição de Combustívela"
    Exit Sub
End Sub
Private Sub cbo_combustivel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cbo_distribuicao.Visible = True Then
            cbo_distribuicao.SetFocus
        Else
            cmd_ok.SetFocus
        End If
    End If
End Sub
Private Sub cbo_combustivel_LostFocus()
    If cbo_combustivel.ListIndex <> -1 Then
        l_tipo_combustivel = Mid(cbo_combustivel.Text, 1, 2)
        PreparaTanques (l_tipo_combustivel)
        cmd_ok.SetFocus
    End If
End Sub
Private Sub cbo_distribuicao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub cmd_data_i_Click()
    g_string = msk_data_i.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i.Text = RetiraGString(1)
        cbo_combustivel.SetFocus
    Else
        msk_data_i.Text = RetiraGString(1)
        cbo_combustivel.SetFocus
    End If
    g_string = " "
End Sub
Private Sub cmd_mais_Click()
Dim x_data As Date
    x_data = msk_data_i.Text
    x_data = x_data + 1
    msk_data_i.Text = Format(x_data, "dd/mm/yyyy")
    l_data_i = x_data
    If ValidaCampos Then
        Consulta
    End If
End Sub
Private Sub cmd_menos_Click()
Dim x_data As Date
    x_data = msk_data_i.Text
    x_data = x_data - 1
    msk_data_i.Text = Format(x_data, "dd/mm/yyyy")
    l_data_i = x_data
    If ValidaCampos Then
        Consulta
    End If
End Sub
Private Sub cmd_ok_Click()
    If ValidaCampos Then
        Consulta
    End If
End Sub
Private Sub cmd_sair_Click()
    Finaliza
    Unload Me
End Sub
Private Sub Form_Activate()
    If Not IsDate(msk_data_i.Text) Then
        If g_lmc = 2 Then
            cbo_distribuicao.ListIndex = 3
        End If
        l_litros_perdas_sobras = 0
        l_processamento = False
    End If
    PreparaLimitadoresPerdasSobras
End Sub
Private Sub PreparaLimitadoresPerdasSobras()
    lLimiteAcumulPerdasSobras = 300 - 1
    If ConfiguracaoDiversa.LocalizarCodigo(1, "LIMITE ACUMULADO PERDAS E SOBRAS") Then
        lLimiteAcumulPerdasSobras = ConfiguracaoDiversa.Valor
    End If

    If lLimiteAcumulPerdasSobras >= 299 Then
        lLimiteAcumulPerdasSobras = 299
        lLimiteDiaPerdasSobras = 199
        lLimiteAcumuladoPerdas = -199
        lLimiteAcumuladoSobras = 199
    ElseIf lLimiteAcumulPerdasSobras >= 249 Then
        lLimiteAcumulPerdasSobras = 249
        lLimiteDiaPerdasSobras = 169
        lLimiteAcumuladoPerdas = -199
        lLimiteAcumuladoSobras = 199
    ElseIf lLimiteAcumulPerdasSobras >= 199 Then
        lLimiteAcumulPerdasSobras = 199
        lLimiteDiaPerdasSobras = 139
        lLimiteAcumuladoPerdas = -199
        lLimiteAcumuladoSobras = 199
    ElseIf lLimiteAcumulPerdasSobras >= 149 Then
        lLimiteAcumulPerdasSobras = 149
        lLimiteDiaPerdasSobras = 99
        lLimiteAcumuladoPerdas = -149
        lLimiteAcumuladoSobras = 149
    ElseIf lLimiteAcumulPerdasSobras >= 99 Then
        lLimiteAcumulPerdasSobras = 99
        lLimiteDiaPerdasSobras = 49
        lLimiteAcumuladoPerdas = -99
        lLimiteAcumuladoSobras = 99
    ElseIf lLimiteAcumulPerdasSobras >= 49 Then
        lLimiteAcumulPerdasSobras = 49
        lLimiteDiaPerdasSobras = 29
        lLimiteAcumuladoPerdas = -49
        lLimiteAcumuladoSobras = 49
    ElseIf lLimiteAcumulPerdasSobras >= 9 Then
        lLimiteAcumulPerdasSobras = 19
        lLimiteDiaPerdasSobras = 9
        lLimiteAcumuladoPerdas = -9
        lLimiteAcumuladoSobras = 9
    End If
End Sub
Function ValidaCampos() As Boolean
    ValidaCampos = False
    If Not IsDate(msk_data_i.Text) Then
        MsgBox "Informe a data inicial.", vbInformation, "Atenção!"
        msk_data_i.SetFocus
    ElseIf cbo_combustivel.ListIndex = -1 Then
        MsgBox "Selecione um combustível.", vbInformation, "Atenção!"
        cbo_combustivel.SetFocus
    ElseIf cbo_distribuicao.ListIndex = -1 And g_lmc = 2 Then
        MsgBox "Selecione o tipo de distribuição.", vbInformation, "Atenção!"
        cbo_distribuicao.SetFocus
    ElseIf Not ValidaLiberacaoLMC Then
        msk_data_i.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Function ValidaLiberacaoLMC() As Boolean
    ValidaLiberacaoLMC = False
    
    If g_nome_usuario = "L.M.C." Then
        If LivroLMC.LocalizarCombustivelConcluido(g_empresa, Mid(cbo_combustivel.Text, 1, 2), CDate(msk_data_i.Text)) = "NAO" Then
            ValidaLiberacaoLMC = True
        ElseIf LivroLMC.LocalizarCombustivelConcluido(g_empresa, Mid(cbo_combustivel.Text, 1, 2), CDate(msk_data_i.Text)) = "SIM" Then
            MsgBox "O LMC está concluído nesta data.", vbCritical, "LMC concluído!"
        ElseIf LivroLMC.LocalizarCombustivelConcluido(g_empresa, Mid(cbo_combustivel.Text, 1, 2), CDate(msk_data_i.Text)) = "**" Then
            MsgBox "O LMC não está com páginas cadastradas.", vbCritical, "Página não cadastrada!"
        End If
    Else
        ValidaLiberacaoLMC = True
    End If
End Function
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
    Screen.MousePointer = 1
    Call GravaAuditoria(1, Me.name, 1, "")
    CentraForm Me
    
    If g_nome_usuario = "L.M.C." Then
        EntradaCombustivel.NomeTabela = "Entrada_Combustivel_LMC"
        MedicaoCombustivel.NomeTabela = "MedicaoCombustivelLMC"
        MovimentoBomba.NomeTabela = "Movimento_Bomba_LMC"
    Else
        EntradaCombustivel.NomeTabela = "Entrada_Combustivel"
        MedicaoCombustivel.NomeTabela = "MedicaoCombustivel"
        MovimentoBomba.NomeTabela = "Movimento_Bomba"
    End If
    If g_lmc <> 2 Then
        lbl_distribuicao.Visible = False
        cbo_distribuicao.Visible = False
    End If
    PreencheCboCombustivel
    PreencheCboTipoDistribuicao
End Sub
Private Sub PreencheCboCombustivel()
    Dim xSQL As String
    Dim rsTabela As New adodb.Recordset
    
    cbo_combustivel.Clear
    'Prepara SQL
    xSQL = "SELECT Nome, Codigo"
    xSQL = xSQL & "  FROM Combustivel"
    xSQL = xSQL & " WHERE Empresa = " & g_empresa
    xSQL = xSQL & " ORDER BY Nome"
    'Abre RecordSet
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
Private Sub PreencheCboTipoDistribuicao()
    cbo_distribuicao.Clear
    cbo_distribuicao.AddItem "Nenhuma"
    cbo_distribuicao.AddItem "+"
    cbo_distribuicao.AddItem "-"
    cbo_distribuicao.AddItem "+ / -"
End Sub
Private Sub PreparaTanques(ByVal pTipoCombustivel As String)
    Dim i As Integer

    For i = 0 To 10
        lCodigoTanque(i) = 0
        lEstoqueTanque(i) = 0
    Next
    
    lSQL = "SELECT [Numero do Tanque] AS Numero FROM Tanque_Combustivel"
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    lSQL = lSQL & " AND [Tipo de Combustivel] = " & preparaTexto(pTipoCombustivel)
    lSQL = lSQL & " ORDER BY [Numero do Tanque]"
    Set rstTanques = Conectar.RsConexao(lSQL)
    lQuantidadeTanque = 0
    '''Call AtualizaRstAfericao(1)
    If rstTanques.RecordCount > 0 Then
        Do Until rstTanques.EOF
            lQuantidadeTanque = lQuantidadeTanque + 1
            lCodigoTanque(lQuantidadeTanque) = rstTanques!numero
            rstTanques.MoveNext
        Loop
    End If
    rstTanques.Close
End Sub
Private Sub msk_data_i_GotFocus()
    If Not IsDate(msk_data_i.Text) Then
        msk_data_i.Text = Format(g_data_def, "dd/mm/yyyy")
    End If
    msk_data_i.SelStart = 0
    If IsDate(msk_data_i.Text) Then
        msk_data_i.SelLength = 2
    Else
        msk_data_i.SelLength = 5
    End If
End Sub
Private Sub msk_data_i_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub msk_data_i_LostFocus()
    If IsDate(msk_data_i.Text) Then
        l_data_i = msk_data_i.Text
        If cbo_combustivel.ListIndex = -1 Then
            cbo_combustivel.SetFocus
        End If
    End If
End Sub
Private Sub txt_perdas_sobras_GotFocus()
    If g_lmc <> 2 Then
        cmd_mais.SetFocus
        Exit Sub
    End If
    If cbo_distribuicao.ListIndex > 0 Then
        txt_perdas_sobras.Text = CalculaPerdasSobras
    End If
End Sub
Private Sub txt_perdas_sobras_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        Altera
        cmd_mais.SetFocus
    'Ctrl + R
    ElseIf KeyAscii = 18 Then
        KeyAscii = 0
        txt_perdas_sobras.Text = CalculaPerdasSobras
    End If
End Sub
