VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0842D103-1E19-101B-9AAF-1A1626551E7C}#1.0#0"; "GRAPH32.OCX"
Begin VB.Form grafico_comparativo_combustivel 
   Caption         =   "Gráficos de Comparação das Vendas de Combustíveis"
   ClientHeight    =   7815
   ClientLeft      =   375
   ClientTop       =   630
   ClientWidth     =   11175
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "grf_comparativo_combustivel.frx":0000
   ScaleHeight     =   7815
   ScaleWidth      =   11175
   Begin VB.ComboBox cbo_ano 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   7140
      Width           =   915
   End
   Begin Threed.SSCommand cmd_imprimir 
      Height          =   975
      Left            =   9000
      TabIndex        =   4
      Top             =   6720
      Width           =   915
      _Version        =   65536
      _ExtentX        =   1614
      _ExtentY        =   1720
      _StockProps     =   78
      Caption         =   "&Imprimir"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "grf_comparativo_combustivel.frx":0446
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
   Begin Threed.SSCommand cmd_grafico 
      Height          =   975
      Left            =   7920
      TabIndex        =   3
      Top             =   6720
      Width           =   915
      _Version        =   65536
      _ExtentX        =   1614
      _ExtentY        =   1720
      _StockProps     =   78
      Caption         =   "&Gráfico"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "grf_comparativo_combustivel.frx":1C80
   End
   Begin Threed.SSCommand cmd_sair 
      Cancel          =   -1  'True
      Height          =   975
      Left            =   10080
      TabIndex        =   5
      Top             =   6720
      Width           =   915
      _Version        =   65536
      _ExtentX        =   1614
      _ExtentY        =   1720
      _StockProps     =   78
      Caption         =   "&Sair"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "grf_comparativo_combustivel.frx":4192
   End
   Begin VB.Label Label7 
      Caption         =   "Ano"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   7140
      Width           =   1155
   End
End
Attribute VB_Name = "grafico_comparativo_combustivel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_litros(1 To 12) As Currency
Dim l_ano As Integer
Dim tbl_movimento_bomba As Table
Private Sub Finaliza()
    tbl_movimento_bomba.Close
End Sub
Private Sub LeDados()
    Dim i As Integer
    Dim x_data As String
    For i = 1 To 12
        l_litros(i) = 0
    Next
    With tbl_movimento_bomba
        If .RecordCount > 0 Then
            x_data = "01/01/" & Format(l_ano, "0000")
            .Seek ">", g_empresa, CDate(x_data), 0, 0
            If Not .NoMatch Then
                Do Until .EOF
                    If Year(!Data) <> Year(x_data) Or !Empresa <> g_empresa Then
                         Exit Do
                    End If
                    i = Month(!Data)
                    l_litros(i) = l_litros(i) + ![Quantidade da Saida]
                    .MoveNext
                Loop
            End If
        End If
    End With
End Sub
Private Sub MontaGraficos()
    Dim i As Integer
    LeDados
    grafico.GraphType = 4
    grafico.PrintStyle = 1
    grafico.GridStyle = 3
    grafico.GraphTitle = g_nome_empresa & " - " & l_ano
    'grafico.LeftTitle = "Litros"
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
Private Sub cbo_ano_GotFocus()
    SendMessageLong cbo_ano.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_ano_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_grafico.SetFocus
    End If
End Sub
Private Sub cmd_grafico_Click()
    Unload Me
    Load Me
    MontaGraficos
    Me.Show 1
End Sub
Private Sub cmd_imprimir_Click()
    grafico.DrawMode = 5
End Sub
Private Sub cmd_sair_Click()
    Finaliza
    Unload Me
End Sub
Private Sub Form_Activate()
    PreencheCboAno
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
Private Sub Form_Load()
    Screen.MousePointer = 1
    CentraForm Me
    Set tbl_movimento_bomba = bd_sgp.OpenTable("Movimento_Bomba")
    tbl_movimento_bomba.Index = "id_data"
End Sub
