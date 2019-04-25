VERSION 5.00
Object = "{0842D103-1E19-101B-9AAF-1A1626551E7C}#1.0#0"; "GRAPH32.OCX"
Begin VB.Form grafico_venda_combustivel_anual 
   Caption         =   "Gráficos de Venda de Combustíveis Anual"
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
   Picture         =   "grf_venda_combustivel_anual.frx":0000
   ScaleHeight     =   7635
   ScaleWidth      =   11175
   Begin VB.ComboBox cbo_combustivel 
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   6720
      Width           =   3435
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   9360
      Picture         =   "grf_venda_combustivel_anual.frx":0446
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Imprime o gráfico da venda de combustível anual."
      Top             =   6660
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   10260
      Picture         =   "grf_venda_combustivel_anual.frx":1720
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   6660
      Width           =   795
   End
   Begin VB.CommandButton cmd_grafico 
      Caption         =   "&Gráfico"
      Height          =   855
      Left            =   8460
      Picture         =   "grf_venda_combustivel_anual.frx":29FA
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Prepara o gráfico da venda de combustível anual."
      Top             =   6660
      Width           =   795
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
      Caption         =   "&Combustível"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   6720
      Width           =   1515
   End
   Begin VB.Label Label7 
      Caption         =   "Ano"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   7140
      Width           =   1515
   End
End
Attribute VB_Name = "grafico_venda_combustivel_anual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_litros(1 To 12) As Currency
Dim l_ano As Integer
Dim lTipoCombustivel As String
Dim lNomeCombustivel As String

Private Combustivel As New cCombustivel
Private MovimentoAfericao As New cMovimentoAfericao
Private MovimentoBomba As New cMovimentoBomba
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
End Sub
Private Sub LeDados()
    Dim i As Integer
    Dim xData As String
    Dim xDataI As Date
    Dim xDataF As Date
    Dim xTipoCombustivel As String
    
    If l_ano = 0 Then
        l_ano = Year(Date)
    End If
    xTipoCombustivel = lTipoCombustivel
    If xTipoCombustivel = "**" Then
        xTipoCombustivel = ""
    End If
    
    For i = 1 To 12
        xData = "01/" & Format(i, "00") & "/" & Format(l_ano, "0000")
        xDataI = CDate(xData)
        Mid(xData, 1, 2) = 32
        Do Until IsDate(xData)
            Mid(xData, 1, 2) = Val(Mid(xData, 1, 2)) - 1
        Loop
        xDataF = CDate(xData)
    
    
        l_litros(i) = MovimentoBomba.QuantidadeVendaData(g_empresa, xDataI, xDataF, xTipoCombustivel, 0)
        l_litros(i) = l_litros(i) - MovimentoAfericao.TotalQtdPeriodoCombustivel(g_empresa, xDataI, xDataF, 1, 9, xTipoCombustivel, "")
    Next
'    With tbl_movimento_bomba
'        If .RecordCount > 0 Then
'            x_data = "01/01/" & Format(l_ano, "0000")
'            .Seek ">", g_empresa, CDate(x_data), 0, 0
'            If Not .NoMatch Then
'                Do Until .EOF
'                    If Year(!Data) <> Year(x_data) Or !Empresa <> g_empresa Then
'                         Exit Do
'                    End If
'                    i = Month(!Data)
'                    If ![Tipo de Combustivel] = lTipoCombustivel Or lTipoCombustivel = "**" Then
'                        l_litros(i) = l_litros(i) + ![Quantidade da Saida]
'                    End If
'                    .MoveNext
'                Loop
'            End If
'        End If
'    End With
End Sub
Private Sub MontaGraficos()
    Dim i As Integer
    LeDados
    grafico.GraphType = 4
    grafico.PrintStyle = 1
    grafico.GridStyle = 3
    grafico.GraphTitle = g_nome_empresa & " - " & l_ano
    grafico.LeftTitle = lNomeCombustivel
    For i = 1 To 12
        grafico.LegendText = Format(i, "00") & "/" & Format(l_ano, "00")
    Next
    For i = 1 To 12
        grafico.LabelText = "  " & l_litros(i) & " Lt"
    Next
    For i = 1 To 12
        grafico.ColorData = i
    Next
    For i = 1 To 12
        grafico.GraphData = l_litros(i)
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
Private Sub cbo_ano_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_grafico.SetFocus
    End If
End Sub
Private Sub cbo_combustivel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_ano.SetFocus
    End If
End Sub
Private Sub cbo_combustivel_LostFocus()
    Dim i As Integer
    If cbo_combustivel.ListIndex <> -1 Then
        lTipoCombustivel = Mid(cbo_combustivel, 1, 2)
        If lTipoCombustivel = "**" Then
            lNomeCombustivel = "Todos os Combustíveis"
        Else
            i = Len(cbo_combustivel.Text)
            lNomeCombustivel = Mid(cbo_combustivel.Text, 6, i - 5)
            If Combustivel.LocalizarCodigo(g_empresa, lTipoCombustivel) Then
            Else
                MsgBox "Combustível inexistente", vbInformation, "Erro Interidade!"
                cbo_combustivel.SetFocus
                Exit Sub
            End If
        End If
    Else
        cbo_combustivel.SetFocus
    End If
End Sub
Private Sub cmd_grafico_Click()
    If ValidaCampos Then
        Unload Me
        Load Me
        MontaGraficos
        Me.Show '1
    End If
End Sub
Private Sub cmd_imprimir_Click()
    grafico.DrawMode = 5
End Sub
Private Sub cmd_sair_Click()
    Finaliza
    Unload Me
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    PreencheCboAno
    PreencheCboCombustivel
    MontaGraficos
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
Private Sub PreencheCboCombustivel()
    Dim rstCombustivel As New adodb.Recordset
    Dim i As Integer
    
    cbo_combustivel.Clear
    cbo_combustivel.AddItem "** - Todos"
    Set rstCombustivel = Conectar.RsConexao("SELECT Codigo, Nome FROM Combustivel WHERE Empresa = " & g_empresa & " ORDER BY Nome")
    'loop RecordSet
    With rstCombustivel
        If Not .BOF Or Not .EOF Then
            .MoveFirst
            Do Until .EOF
                cbo_combustivel.AddItem !Codigo & " - " & !Nome
                .MoveNext
            Loop
        End If
        .Close
    End With
    Set rstCombustivel = Nothing

    For i = 0 To cbo_combustivel.ListCount - 1
        cbo_combustivel.ListIndex = i
        If Mid(cbo_combustivel.Text, 1, 2) = lTipoCombustivel Then
            cbo_combustivel.ListIndex = i
            Exit For
        End If
        cbo_combustivel.ListIndex = -1
    Next
End Sub
Function ValidaCampos() As Boolean
    ValidaCampos = False
    If cbo_ano.ListIndex = -1 Then
        MsgBox "Selecione o ano.", vbInformation, "Atenção!"
        cbo_ano.SetFocus
    ElseIf cbo_combustivel.ListIndex = -1 Then
        MsgBox "Selecione o combustível.", vbInformation, "Atenção!"
        cbo_combustivel.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub Form_Load()
    Screen.MousePointer = 1
    CentraForm Me
    If g_nome_usuario = "L.M.C." Then
        MovimentoAfericao.NomeTabela = "Movimento_Afericao_LMC"
        MovimentoBomba.NomeTabela = "Movimento_Bomba_LMC"
    Else
        MovimentoAfericao.NomeTabela = "Movimento_Afericao"
        MovimentoBomba.NomeTabela = "Movimento_Bomba"
    End If
End Sub
