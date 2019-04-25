VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0842D103-1E19-101B-9AAF-1A1626551E7C}#1.0#0"; "GRAPH32.OCX"
Begin VB.Form grafico_venda_combustivel 
   Caption         =   "Gráficos de Vendas de Combustíveis"
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
   Picture         =   "grf_venda_combustivel.frx":0000
   ScaleHeight     =   7815
   ScaleWidth      =   11175
   Begin VB.ComboBox cbo_vendas 
      Height          =   300
      Left            =   2220
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   6720
      Width           =   1995
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   615
      Left            =   180
      TabIndex        =   3
      Top             =   7080
      Width           =   4035
      _Version        =   65536
      _ExtentX        =   7117
      _ExtentY        =   1085
      _StockProps     =   14
      Caption         =   "Tipo de Gráfico"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSOption opt_grafico 
         Height          =   195
         Index           =   1
         Left            =   2220
         TabIndex        =   5
         Top             =   300
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "Pizza 3D"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption opt_grafico 
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   4
         Top             =   300
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "Barra 3D"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin Threed.SSCommand cmd_imprimir 
      Height          =   975
      Left            =   9000
      TabIndex        =   7
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
      Picture         =   "grf_venda_combustivel.frx":0446
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
   Begin Threed.SSCommand cmd_grafico 
      Height          =   975
      Left            =   7920
      TabIndex        =   6
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
      Picture         =   "grf_venda_combustivel.frx":1C80
   End
   Begin Threed.SSCommand cmd_sair 
      Cancel          =   -1  'True
      Height          =   975
      Left            =   10080
      TabIndex        =   8
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
      Picture         =   "grf_venda_combustivel.frx":4192
   End
   Begin VB.Label Label2 
      Caption         =   "Vendas de Combustíveis"
      Height          =   255
      Left            =   180
      TabIndex        =   1
      Top             =   6720
      Width           =   1935
   End
End
Attribute VB_Name = "grafico_venda_combustivel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_litros_1 As Currency
Dim l_litros_2 As Currency
Dim l_litros_3 As Currency
Dim l_litros_4 As Currency
Dim l_litros_5 As Currency
Dim l_litros_6 As Currency
Dim l_tipo_grafico As Integer
Dim tbl_acm_venda_combustivel As Table
Private Sub Finaliza()
    tbl_acm_venda_combustivel.Close
End Sub
Private Sub MontaGraficos()
    If l_tipo_grafico = 0 Then
        grafico.GraphType = 4
    Else
        grafico.GraphType = 2
    End If
    grafico.PrintStyle = 1
    grafico.GridStyle = 3
    grafico.GraphTitle = "Vendas de Combustíveis " & Format(g_data, "mm") & "/" & Year(g_data)
    'grafico.LeftTitle = "Litros"
    
    grafico.LegendText = "PL"
    grafico.LegendText = "87"
    grafico.LegendText = "TK"
    grafico.LegendText = "CO"
    grafico.LegendText = "GO"
    grafico.LegendText = ""
    grafico.LabelText = "  " & l_litros_1 & " Lt"
    grafico.LabelText = "  " & l_litros_2 & " Lt"
    grafico.LabelText = "  " & l_litros_3 & " Lt"
    grafico.LabelText = "  " & l_litros_4 & " Lt"
    grafico.LabelText = "  " & l_litros_5 & " Lt"
    grafico.LabelText = "  " & l_litros_6 & " Lt"
    grafico.ColorData = 1
    grafico.ColorData = 2
    grafico.ColorData = 3
    grafico.ColorData = 4
    grafico.ColorData = 5
    grafico.ColorData = 6
    grafico.GraphData = l_litros_1
    grafico.GraphData = l_litros_2
    grafico.GraphData = l_litros_3
    grafico.GraphData = l_litros_4
    grafico.GraphData = l_litros_5
    grafico.GraphData = l_litros_6
    'grafico.ThisPoint = 4
    grafico.ExtraData = 7
    grafico.ExtraData = 7
    grafico.ExtraData = 7
    grafico.ExtraData = 7
    grafico.ExtraData = 7
    grafico.ExtraData = 7
End Sub
Private Sub cbo_vendas_Click()
    If Not (tbl_acm_venda_combustivel.BOF And tbl_acm_venda_combustivel.EOF) Then
        If cbo_vendas.ListIndex <> -1 Then
            g_data = "01/" & cbo_vendas.Text
        Else
            cbo_vendas.SetFocus
        End If
    End If
End Sub
Private Sub cbo_vendas_GotFocus()
    SendMessageLong cbo_vendas.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_vendas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{Tab}"
    End If
End Sub
Private Sub cmd_grafico_Click()
    If opt_grafico(0) Then
        l_tipo_grafico = 0
    Else
        l_tipo_grafico = 1
    End If
    If tbl_acm_venda_combustivel.RecordCount > 0 Then
        Unload Me
        Load Me
        tbl_acm_venda_combustivel.Seek "=", g_data
        l_litros_1 = tbl_acm_venda_combustivel!litros_2
        l_litros_2 = tbl_acm_venda_combustivel!litros_3
        l_litros_3 = tbl_acm_venda_combustivel!litros_4
        l_litros_4 = tbl_acm_venda_combustivel!litros_6
        l_litros_5 = tbl_acm_venda_combustivel!litros_9
        l_litros_6 = tbl_acm_venda_combustivel!litros_1
        MontaGraficos
        Me.Show
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
    If Not (tbl_acm_venda_combustivel.BOF And tbl_acm_venda_combustivel.EOF) Then
        PreencheCboVendas
    End If
End Sub
Private Sub PreencheCboVendas()
    cbo_vendas.Clear
    tbl_acm_venda_combustivel.MoveFirst
    Do Until tbl_acm_venda_combustivel.EOF
        cbo_vendas.AddItem Format(Month(tbl_acm_venda_combustivel!mes_ano), "00") & "/" & Mid(Year(tbl_acm_venda_combustivel!mes_ano), 3, 2)
        tbl_acm_venda_combustivel.MoveNext
    Loop
End Sub
Private Sub Form_Load()
    Screen.MousePointer = 1
    CentraForm Me
    Set tbl_acm_venda_combustivel = bd_sgp.OpenTable("acm_venda_combustiveis")
    tbl_acm_venda_combustivel.Index = "id_mes_ano"
    PreencheCboVendas
End Sub
Private Sub opt_grafico_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{Tab}"
        cmd_grafico_Click
    End If
End Sub
