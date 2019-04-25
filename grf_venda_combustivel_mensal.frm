VERSION 5.00
Object = "{0842D103-1E19-101B-9AAF-1A1626551E7C}#1.0#0"; "GRAPH32.OCX"
Begin VB.Form grafico_venda_combustivel_mensal 
   Caption         =   "Gráficos de Vendas de Combustíveis Mensal"
   ClientHeight    =   7815
   ClientLeft      =   450
   ClientTop       =   570
   ClientWidth     =   11175
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "grf_venda_combustivel_mensal.frx":0000
   ScaleHeight     =   7815
   ScaleWidth      =   11175
   Begin VB.CommandButton cmd_grafico 
      Caption         =   "&Gráfico"
      Height          =   855
      Left            =   8460
      Picture         =   "grf_venda_combustivel_mensal.frx":0446
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Prepara o gráfico da venda de combustível mensal."
      Top             =   6840
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   10260
      Picture         =   "grf_venda_combustivel_mensal.frx":1720
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   6840
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   9360
      Picture         =   "grf_venda_combustivel_mensal.frx":29FA
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Imprime o gráfico da venda de combustível mensal."
      Top             =   6840
      Width           =   795
   End
   Begin VB.ComboBox cbo_mes_ano 
      Height          =   300
      Left            =   2220
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   6720
      Width           =   1995
   End
   Begin VB.Frame frmTipoGrafico 
      Caption         =   "Tipo de Gráfico"
      Height          =   615
      Left            =   180
      TabIndex        =   3
      Top             =   7080
      Width           =   4035
      Begin VB.OptionButton optGrafico 
         Caption         =   "Pizza 3D"
         Height          =   195
         Index           =   1
         Left            =   2220
         TabIndex        =   5
         Top             =   300
         Width           =   1575
      End
      Begin VB.OptionButton optGrafico 
         Caption         =   "Barra 3D"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   4
         Top             =   300
         Width           =   1575
      End
   End
   Begin GraphLib.Graph grafico 
      Height          =   6435
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10935
      _Version        =   65536
      _ExtentX        =   19288
      _ExtentY        =   11351
      _StockProps     =   96
      BorderStyle     =   1
      GraphType       =   4
      NumPoints       =   6
      PrintStyle      =   3
      RandomData      =   1
      ColorData       =   0
      ExtraData       =   0
      ExtraData[]     =   0
      FontFamily      =   4
      FontSize        =   4
      FontSize[0]     =   200
      FontSize[1]     =   150
      FontSize[2]     =   100
      FontSize[3]     =   100
      FontStyle       =   4
      GraphData       =   0
      GraphData[]     =   0
      LabelText       =   5
      LabelText[0]    =   "1"
      LabelText[1]    =   "2"
      LabelText[2]    =   "3"
      LabelText[3]    =   "4"
      LabelText[4]    =   "5"
      LegendText      =   6
      LegendText[0]   =   "NS"
      LegendText[1]   =   "PL"
      LegendText[2]   =   "87"
      LegendText[3]   =   "TK"
      LegendText[4]   =   "68"
      LegendText[5]   =   "CO"
      PatternData     =   0
      SymbolData      =   0
      XPosData        =   0
      XPosData[]      =   0
   End
   Begin VB.Label Label2 
      Caption         =   "Mês / Ano"
      Height          =   255
      Left            =   180
      TabIndex        =   1
      Top             =   6720
      Width           =   1935
   End
End
Attribute VB_Name = "grafico_venda_combustivel_mensal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lLitro(1 To 12) As Currency
Dim l_mes As Integer
Dim l_ano As Integer
Dim l_qtd_empresa As Integer
Dim l_numero_empresa(1 To 12) As Integer
Dim l_nome_empresa(1 To 12) As String
Dim l_titulo_empresa As String
Dim l_tipo_grafico As Integer

Private MovimentoAfericao As New cMovimentoAfericao
Private MovimentoBomba As New cMovimentoBomba
Private Sub Finaliza()
    Set MovimentoAfericao = Nothing
    Set MovimentoBomba = Nothing
End Sub
Private Sub LeDados()
    Dim i As Integer
    Dim xData As String
    Dim rsEmpresa As New adodb.Recordset
    Dim xDataInicial As Date
    Dim xDataFinal As Date

    'Pega Nome do Grupo de Empresas e as Abreviações das empresas do grupo
    l_titulo_empresa = ReadINI("GRUPO DE EMPRESAS", "Nome do Grupo", gArquivoIni)
    'Pega Nome as Abreviações das Empresas do grupo
    g_string = ReadINI("GRUPO DE EMPRESAS", "Nome Abreviado das Empresas", gArquivoIni)
    For i = 1 To 12
        lLitro(i) = 0
        l_numero_empresa(i) = 0
        l_nome_empresa(i) = RetiraGString(i)
    Next
    g_string = ""
        
        
    Set rsEmpresa = Conectar.RsConexao("SELECT Codigo FROM Empresas WHERE Inativo = " & preparaBooleano(False) & " ORDER BY Codigo")
    'loop RecordSet
    l_qtd_empresa = 0
    With rsEmpresa
        If .RecordCount > 0 Then
            .MoveFirst
            Do Until .EOF
                l_qtd_empresa = l_qtd_empresa + 1
                l_numero_empresa(l_qtd_empresa) = !Codigo
                .MoveNext
            Loop
        End If
        .Close
    End With
    Set rsEmpresa = Nothing
        
    xDataInicial = CDate("01/" & Format(l_mes, "00") & "/" & Format(l_ano, "0000"))
    i = 32
    xData = i & "/" & Format(l_mes, "00") & "/" & Format(l_ano, "0000")
    Do Until IsDate(xData)
        i = i - 1
        Mid(xData, 1, 2) = i
    Loop
    xDataFinal = CDate(xData)
    
    For i = 1 To l_qtd_empresa
        lLitro(i) = MovimentoBomba.TotalLitrosData(l_numero_empresa(i), xDataInicial, xDataFinal, "", "Movimento_Bomba")
        lLitro(i) = lLitro(i) - MovimentoAfericao.TotalQtdPeriodoCombustivel(l_numero_empresa(i), xDataInicial, xDataFinal, 1, 9, "", "")
    Next
End Sub
Private Sub MontaGraficos()
    Dim i As Integer
    Dim x_data As Date
    x_data = "01/" & l_mes & "/" & l_ano
    LeDados
    If l_tipo_grafico = 0 Then
        grafico.GraphType = 4
    Else
        grafico.GraphType = 2
    End If
    grafico.PrintStyle = 1
    grafico.GridStyle = 3
    If l_qtd_empresa = 1 Then
        l_qtd_empresa = 2
    End If
    grafico.NumPoints = l_qtd_empresa
    grafico.GraphTitle = l_titulo_empresa & " - " & Format(x_data, "mmmm") & " de " & l_ano
    grafico.LeftTitle = "Combustível"
    For i = 1 To l_qtd_empresa
        grafico.LegendText = l_nome_empresa(i)
    Next
    For i = 1 To l_qtd_empresa
        grafico.LabelText = Format(lLitro(i), "###,###,##0") & " Lt"
    Next
    For i = 1 To l_qtd_empresa
        grafico.ColorData = i
    Next
    For i = 1 To l_qtd_empresa
        grafico.GraphData = lLitro(i)
    Next
    'grafico.ThisPoint = 4
    For i = 1 To l_qtd_empresa
        grafico.ExtraData = 7
    Next
End Sub
Private Sub cbo_mes_ano_Click()
    If cbo_mes_ano.ListIndex <> -1 Then
        l_mes = Mid(cbo_mes_ano.Text, 1, 2)
        l_ano = Mid(cbo_mes_ano.Text, 4, 4)
    Else
        cbo_mes_ano.SetFocus
    End If
End Sub
Private Sub cbo_mes_ano_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        optGrafico(0).SetFocus
    End If
End Sub
Private Sub cmd_grafico_Click()
    If optGrafico(0) Then
        l_tipo_grafico = 0
    Else
        l_tipo_grafico = 1
    End If
    Unload Me
    Load Me
    MontaGraficos
    Me.Show
End Sub
Private Sub cmd_imprimir_Click()
    grafico.DrawMode = 5
End Sub
Private Sub cmd_sair_Click()
    Finaliza
    Unload Me
End Sub
Private Sub Form_Activate()
    Dim i As Integer
    Dim x_mes As Integer
    Dim x_ano As Integer
    PreencheCboMesAno
    If Val(l_ano) > 0 Then
        x_mes = l_mes
        x_ano = l_ano
        For i = 0 To cbo_mes_ano.ListCount - 1
            cbo_mes_ano.ListIndex = i
            If Mid(cbo_mes_ano.Text, 1, 2) = x_mes And Mid(cbo_mes_ano.Text, 4, 4) = x_ano Then
                Exit For
            End If
        Next
    Else
        cbo_mes_ano.ListIndex = cbo_mes_ano.ListCount - 1
    End If
    MontaGraficos
End Sub
Private Sub PreencheCboMesAno()
    Dim i As Integer
    Dim x_ano As Integer
    Dim x_mes As Integer
    cbo_mes_ano.Clear
    For x_ano = 1995 To Year(g_data_def)
        For x_mes = 1 To 12
            cbo_mes_ano.AddItem Format(x_mes, "00") & "/" & Format(x_ano, "0000")
            cbo_mes_ano.ItemData(cbo_mes_ano.NewIndex) = Format(x_mes, "00") & Format(x_ano, "0000")
        Next
    Next
End Sub
Private Sub Form_Load()
    Screen.MousePointer = 1
    CentraForm Me
    MovimentoAfericao.NomeTabela = "Movimento_Afericao"
    MovimentoBomba.NomeTabela = "Movimento_Bomba"
End Sub
Private Sub optGrafico_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_grafico.SetFocus
    End If
End Sub
