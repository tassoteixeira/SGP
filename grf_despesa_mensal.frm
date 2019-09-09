VERSION 5.00
Object = "{0842D103-1E19-101B-9AAF-1A1626551E7C}#1.0#0"; "GRAPH32.OCX"
Begin VB.Form grafico_despesa_mensal 
   Caption         =   "Gráficos de Despesas Mensal"
   ClientHeight    =   7635
   ClientLeft      =   375
   ClientTop       =   630
   ClientWidth     =   11175
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "grf_despesa_mensal.frx":0000
   ScaleHeight     =   7635
   ScaleWidth      =   11175
   Begin VB.CommandButton cmd_grafico 
      Caption         =   "&Gráfico"
      Height          =   855
      Left            =   8460
      Picture         =   "grf_despesa_mensal.frx":0446
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Prepara o gráfico de despesa mensal."
      Top             =   6660
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   10260
      Picture         =   "grf_despesa_mensal.frx":1720
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   6660
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   9360
      Picture         =   "grf_despesa_mensal.frx":29FA
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Imprime o gráfico de despesa mensal."
      Top             =   6660
      Width           =   795
   End
   Begin VB.ComboBox cbo_conta 
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   6720
      Width           =   3435
   End
   Begin VB.ComboBox cbo_mes_ano 
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   7140
      Width           =   1155
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
      NumPoints       =   12
      PrintStyle      =   1
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
      LegendText      =   5
      LegendText[0]   =   "NS"
      LegendText[1]   =   "PL"
      LegendText[2]   =   "87"
      LegendText[3]   =   "TK"
      LegendText[4]   =   "68"
      PatternData     =   0
      SymbolData      =   0
      XPosData        =   0
      XPosData[]      =   0
   End
   Begin VB.Label Label3 
      Caption         =   "&Conta"
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   6720
      Width           =   1515
   End
   Begin VB.Label Label7 
      Caption         =   "Mês/Ano"
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   7140
      Width           =   1515
   End
End
Attribute VB_Name = "grafico_despesa_mensal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lTotal(1 To 12) As Currency
Dim l_mes As Integer
Dim l_ano As Integer
Dim l_conta As Integer
Dim l_nome_conta As String
Dim l_qtd_empresa As Integer
Dim l_numero_empresa(1 To 12) As Integer
Dim l_nome_empresa(1 To 12) As String * 20
Dim l_titulo_empresa As String

Private BaixaPagar As New cBaixaPagar
Private Conta As New cContas
Private Sub Finaliza()
    Set BaixaPagar = Nothing
    Set Conta = Nothing
End Sub
Private Sub LeDados()
    Dim i As Integer
    Dim xData As String
    Dim rsEmpresa As New ADODB.Recordset
    Dim xDataInicial As Date
    Dim xDataFinal As Date
    
    'Pega Nome do Grupo de Empresas e as Abreviações das empresas do grupo
    l_titulo_empresa = ReadINI("GRUPO DE EMPRESAS", "Nome do Grupo", gArquivoIni)
    'Pega Nome as Abreviações das Empresas do grupo
    g_string = ReadINI("GRUPO DE EMPRESAS", "Nome Abreviado das Empresas", gArquivoIni)
    For i = 1 To 12
        lTotal(i) = 0
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
        lTotal(i) = BaixaPagar.TotalPeriodoConta(l_numero_empresa(i), xDataInicial, xDataFinal, l_conta)
    Next
End Sub
Private Sub MontaGraficos()
    Dim i As Integer
    Dim x_data As Date
    x_data = "01/" & l_mes & "/" & l_ano
    LeDados
    grafico.GraphType = 4
    grafico.PrintStyle = 1
    grafico.GridStyle = 3
    If l_qtd_empresa = 1 Then
        l_qtd_empresa = 2
    End If
    grafico.NumPoints = l_qtd_empresa
    grafico.GraphTitle = l_titulo_empresa & " - " & Format(x_data, "mmmm") & " de " & l_ano
    grafico.LeftTitle = l_nome_conta
    For i = 1 To l_qtd_empresa
        grafico.LegendText = l_nome_empresa(i)
    Next
    For i = 1 To l_qtd_empresa
        grafico.LabelText = Format(lTotal(i), "###,###,##0.00")
    Next
    For i = 1 To l_qtd_empresa
        grafico.ColorData = i
    Next
    For i = 1 To l_qtd_empresa
        grafico.GraphData = lTotal(i)
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
        cmd_grafico_Click
    End If
End Sub
Private Sub cbo_conta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_mes_ano.SetFocus
    End If
End Sub
Private Sub cbo_conta_LostFocus()
    l_conta = 0
    If cbo_conta.ListIndex <> -1 Then
        l_conta = cbo_conta.ItemData(cbo_conta.ListIndex)
        l_nome_conta = cbo_conta
    End If
End Sub
Private Sub cmd_grafico_Click()
    Unload Me
    Load Me
    MontaGraficos
    Me.Show '1
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
    Dim x_conta As Integer
    PreencheCboConta
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
    If Val(l_conta) > 0 Then
        x_conta = l_conta
        For i = 0 To cbo_conta.ListCount - 1
            cbo_conta.ListIndex = i
            If cbo_conta.ItemData(cbo_conta.ListIndex) = x_conta Then
                Exit For
            End If
        Next
    Else
        cbo_conta.ListIndex = 0
    End If
    MontaGraficos
End Sub
Private Sub Form_Load()
    Screen.MousePointer = 1
    CentraForm Me
End Sub
Private Sub PreencheCboConta()
    Dim rsConta As New ADODB.Recordset
    
    cbo_conta.Clear
    cbo_conta.AddItem "Todas as Contas"
    cbo_conta.ItemData(cbo_conta.NewIndex) = 0
    
    Set rsConta = Conectar.RsConexao("SELECT Codigo, Nome FROM Contas WHERE Empresa = " & g_empresa & " ORDER BY Nome")
    'loop RecordSet
    With rsConta
        If .RecordCount > 0 Then
            .MoveFirst
            Do Until .EOF
                cbo_conta.AddItem !Nome
                cbo_conta.ItemData(cbo_conta.NewIndex) = !Codigo
                .MoveNext
            Loop
        End If
        .Close
    End With
    Set rsConta = Nothing
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
'    cbo_mes_ano.ListIndex = cbo_mes_ano.ListCount
End Sub
