VERSION 5.00
Object = "{0842D103-1E19-101B-9AAF-1A1626551E7C}#1.0#0"; "GRAPH32.OCX"
Begin VB.Form grafico_despesa_anual 
   Caption         =   "Gráficos de Despesas Anual"
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
   Picture         =   "grf_despesa_anual.frx":0000
   ScaleHeight     =   7635
   ScaleWidth      =   11175
   Begin VB.CommandButton cmd_grafico 
      Caption         =   "&Gráfico"
      Height          =   855
      Left            =   8460
      Picture         =   "grf_despesa_anual.frx":0446
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Prepara o gráfico de despesa anual."
      Top             =   6660
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   10260
      Picture         =   "grf_despesa_anual.frx":1720
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
      Picture         =   "grf_despesa_anual.frx":29FA
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Imprime o gráfico de despesa anual."
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
   Begin VB.ComboBox cbo_ano 
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   7140
      Width           =   915
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
      Caption         =   "Ano"
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   7140
      Width           =   1515
   End
End
Attribute VB_Name = "grafico_despesa_anual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lTotal(1 To 12) As Currency
Dim l_ano As Integer
Dim l_conta As Integer
Dim l_nome_conta As String

Private Conta As New cContas
Private BaixaPagar As New cBaixaPagar
Private Sub Finaliza()
    Set BaixaPagar = Nothing
    Set Conta = Nothing
End Sub
Private Sub LeDados()
    Dim xDia As Integer
    Dim xMes As Integer
    Dim xData As String
    Dim xDataInicial As Date
    Dim xDataFinal As Date
    
    For xMes = 1 To 12
        xDataInicial = CDate("01/" & Format(xMes, "00") & "/" & Format(l_ano, "0000"))
        xDia = 32
        xData = xDataInicial
        Mid(xData, 1, 2) = xDia
        Do Until IsDate(xData)
            xDia = xDia - 1
            Mid(xData, 1, 2) = xDia
        Loop
        xDataFinal = CDate(xData)
        lTotal(xMes) = BaixaPagar.TotalPeriodoConta(g_empresa, xDataInicial, xDataFinal, l_conta)
    Next
End Sub
Private Sub MontaGraficos()
    Dim i As Integer
    LeDados
    grafico.GraphType = 4
    grafico.PrintStyle = 1
    grafico.GridStyle = 3
    grafico.GraphTitle = g_nome_empresa & " - " & l_ano
    grafico.LeftTitle = l_nome_conta
    For i = 1 To 12
        grafico.LegendText = Format(i, "00") & "/" & Format(l_ano, "00")
    Next
    For i = 1 To 12
        grafico.LabelText = Format(lTotal(i), "###,###,##0.00")
    Next
    For i = 1 To 12
        grafico.ColorData = i
    Next
    For i = 1 To 12
        grafico.GraphData = lTotal(i)
    Next
    'grafico.ThisPoint = 4
    For i = 1 To 12
        grafico.ExtraData = 7
    Next
End Sub
Private Sub cbo_ano_Click()
    If cbo_ano.ListIndex <> -1 Then
        l_ano = cbo_ano.Text
    Else
        cbo_ano.SetFocus
    End If
End Sub
Private Sub cbo_ano_GotFocus()
    SendMessageLong cbo_ano.hWnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_ano_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_grafico_Click
    End If
End Sub
Private Sub cbo_conta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_ano.SetFocus
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
    Dim x_ano As Integer
    Dim x_conta As Integer
    PreencheCboConta
    PreencheCboAno
    If Val(l_ano) > 0 Then
        x_ano = l_ano
        For i = 0 To cbo_ano.ListCount - 1
            cbo_ano.ListIndex = i
            If cbo_ano = x_ano Then
                Exit For
            End If
        Next
    Else
        cbo_ano.ListIndex = cbo_ano.ListCount - 1
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
Private Sub PreencheCboAno()
    Dim i As Integer
    cbo_ano.Clear
    For i = 1995 To Year(g_data_def)
        cbo_ano.AddItem i
        cbo_ano.ItemData(cbo_ano.NewIndex) = i
    Next
    For i = 0 To cbo_ano.ListCount - 1
        If cbo_ano.ItemData(i) = l_ano Then
            cbo_ano.ListIndex = i
            Exit For
        End If
    Next
End Sub
