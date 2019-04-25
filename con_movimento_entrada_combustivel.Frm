VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form consulta_movimento_entrada_combustivel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta Movimento de Entrada de Combustível"
   ClientHeight    =   5355
   ClientLeft      =   1155
   ClientTop       =   2490
   ClientWidth     =   9105
   Icon            =   "con_movimento_entrada_combustivel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "con_movimento_entrada_combustivel.frx":0442
   ScaleHeight     =   5355
   ScaleWidth      =   9105
   Begin VB.Frame Frame1 
      Caption         =   "S&elecionar por..."
      Height          =   1155
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   6855
      Begin VB.ComboBox cbo_campo 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   300
         Width           =   5775
      End
      Begin VB.TextBox txt_condicao 
         Height          =   285
         Left            =   3120
         TabIndex        =   6
         Top             =   720
         Width           =   3615
      End
      Begin VB.ComboBox cbo_operador 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "&Campo"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   675
      End
      Begin VB.Label Label3 
         Caption         =   "Co&ndição"
         Height          =   255
         Left            =   2340
         TabIndex        =   5
         Top             =   780
         Width           =   795
      End
      Begin VB.Label Label4 
         Caption         =   "O&perador"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   780
         Width           =   735
      End
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   7440
      Picture         =   "con_movimento_entrada_combustivel.frx":0488
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Inicia a pesquisa selecionada."
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   8280
      Picture         =   "con_movimento_entrada_combustivel.frx":1A92
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   120
      Width           =   735
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4035
      Left            =   60
      TabIndex        =   9
      Top             =   1200
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   7117
      _Version        =   393216
      FixedCols       =   0
      BackColorSel    =   12582912
      ForeColorSel    =   16777215
      BackColorBkg    =   12632256
      SelectionMode   =   1
   End
End
Attribute VB_Name = "consulta_movimento_entrada_combustivel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim lFlagConsulta As Integer
    Dim rsDados As New adodb.Recordset
    Dim fld As adodb.Field
    Dim lSQL As String
    Dim lNomeTabela As String
Private Sub AtualizaGrid()

    On Error GoTo ErroConsulta
    
    If ValidaCampos Then
        lSQL = ""
        lSQL = lSQL & "SELECT " & lNomeTabela & ".Data, " & lNomeTabela & ".[Tipo de Combustivel], " & lNomeTabela & ".[Codigo do Fornecedor], Fornecedor.Nome, " & lNomeTabela & ".[Numero da Nota], " & lNomeTabela & ".Item, " & lNomeTabela & ".[Valor do Litro], " & lNomeTabela & ".Quantidade, " & lNomeTabela & ".[Valor da Entrada], " & lNomeTabela & ".[Tipo de Transporte], " & lNomeTabela & ".Ordem"
        lSQL = lSQL & "  FROM " & lNomeTabela & ", Fornecedor"
        If cbo_campo.ListIndex <> -1 Then
            lSQL = lSQL & " WHERE " & lNomeTabela & ".Empresa = " & g_empresa
            lSQL = lSQL & "   AND " & lNomeTabela & ".Empresa = Fornecedor.empresa"
            lSQL = lSQL & "   AND " & lNomeTabela & ".[Codigo do Fornecedor] = Fornecedor.codigo"
            lSQL = lSQL & "   AND " & lNomeTabela & "." & fConsultaPreparaCondicao(cbo_campo.Text, rsDados.Fields(cbo_campo.Text).Type, cbo_operador.Text, txt_condicao.Text)
        End If
        lSQL = lSQL & " ORDER BY Data, [Numero da Nota], [Tipo de Combustivel]"
        Set rsDados = Conectar.RsConexao(lSQL)
        FormataGrid
        PreencheGrid
    End If
    Exit Sub

ErroConsulta:
    If Err = 3075 Then
        MsgBox "Condição inválida.", vbInformation, "Erro de Consulta"
    Else
        MsgBox Error, vbInformation, "Erro de Consulta"
    End If
    Exit Sub
End Sub
Private Sub Finaliza()
    Set rsDados = Nothing
End Sub
Private Sub FormataGrid()
    Dim i As Integer
    'início das variáveis para traduzir
    'Dim VItem As New ValueItem
    'Dim VItems As ValueItems
    'fim das variáveis para traduzir
    
    MSFlexGrid1.Cols = 11
    MSFlexGrid1.ColWidth(0) = 500
    MSFlexGrid1.Rows = 1
    i = 0
    MSFlexGrid1.TextMatrix(0, i) = "Data"
    MSFlexGrid1.ColAlignment(i) = flexAlignCenterCenter
    MSFlexGrid1.ColWidth(i) = 1000
    i = i + 1
    MSFlexGrid1.TextMatrix(0, i) = "Comb"
    MSFlexGrid1.ColAlignment(i) = flexAlignCenterCenter
    MSFlexGrid1.ColWidth(i) = 500
    i = i + 1
    MSFlexGrid1.TextMatrix(0, i) = "Cód."
    MSFlexGrid1.ColAlignment(i) = flexAlignRightCenter
    MSFlexGrid1.ColWidth(i) = 500
    i = i + 1
    MSFlexGrid1.TextMatrix(0, i) = "Nome do Fornecedor"
    MSFlexGrid1.ColAlignment(i) = flexAlignLeftCenter
    MSFlexGrid1.ColWidth(i) = 2000
    i = i + 1
    MSFlexGrid1.TextMatrix(0, i) = "N. da NF"
    MSFlexGrid1.ColAlignment(i) = flexAlignLeftCenter
    MSFlexGrid1.ColWidth(i) = 800
    i = i + 1
    MSFlexGrid1.TextMatrix(0, i) = "Ítem"
    MSFlexGrid1.ColAlignment(i) = flexAlignLeftCenter
    MSFlexGrid1.ColWidth(i) = 500
    i = i + 1
    MSFlexGrid1.TextMatrix(0, i) = "Valor do Litro"
    MSFlexGrid1.ColAlignment(i) = flexAlignRightCenter
    MSFlexGrid1.ColWidth(i) = 1100
    i = i + 1
    MSFlexGrid1.TextMatrix(0, i) = "Quantidade"
    MSFlexGrid1.ColAlignment(i) = flexAlignRightCenter
    MSFlexGrid1.ColWidth(i) = 1000
    i = i + 1
    MSFlexGrid1.TextMatrix(0, i) = "Valor da Entrada"
    MSFlexGrid1.ColAlignment(i) = flexAlignRightCenter
    MSFlexGrid1.ColWidth(i) = 1300
    i = i + 1
    MSFlexGrid1.TextMatrix(0, i) = "Tipo de Transporte"
    MSFlexGrid1.ColAlignment(i) = flexAlignCenterCenter
    MSFlexGrid1.ColWidth(i) = 1500
    i = i + 1
    MSFlexGrid1.TextMatrix(0, i) = "Ordem"
    MSFlexGrid1.ColAlignment(i) = flexAlignCenterCenter
    MSFlexGrid1.ColWidth(i) = 300
    
    'Início da Tradução
    'Set VItems = MSFlexGrid1.Columns(8).ValueItems
    'VItem.Value = 0
    'VItem.DisplayValue = "Entrega"
    'VItems.Add VItem
    'VItem.Value = 1
    'VItem.DisplayValue = "Retirada"
    'VItems.Add VItem
    'VItems.Translate = True
End Sub
Private Sub MarcaCelulas()
    g_string = ""
    If MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 0) <> "" Then
        g_string = MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 0) & "|@|"
        g_string = g_string & MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 1) & "|@|"
        g_string = g_string & MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 4) & "|@|"
        g_string = g_string & MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 10) & "|@|"
    End If
End Sub
Private Sub PreencheGrid()
    Dim i As Integer
    Dim i2 As Integer
    If Not rsDados.EOF Then
        rsDados.MoveFirst
        Do Until rsDados.EOF
            MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
            MSFlexGrid1.Row = MSFlexGrid1.Rows - 1
            i2 = -1
            For Each fld In rsDados.Fields
                i2 = i2 + 1
                MSFlexGrid1.Col = i2
                'If fld.Name = "CODIGO" Then
                '    Grid1.Text = fMascaraContaContabil(fld.Value)
                'Else
                    If IsNull(fld.Value) Then
                        MSFlexGrid1.Text = ""
                    Else
                        If fld.Type = adCurrency Then
                            MSFlexGrid1.Text = Format(fld.Value, "###,###,##0.00")
                        Else
                            MSFlexGrid1.Text = fld.Value
                        End If
                    End If
                'End If
            Next
            rsDados.MoveNext
        Loop
        MSFlexGrid1.Row = 1
        MSFlexGrid1.Col = 0
    End If
End Sub
Private Sub PreencheCampos()
    lSQL = "SELECT * FROM " & lNomeTabela & " WHERE Empresa = 0"
    Set rsDados = Conectar.RsConexao(lSQL)
    cbo_campo.Clear
    For Each fld In rsDados.Fields
        cbo_campo.AddItem fld.name
    Next
End Sub
Private Sub cbo_campo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_operador.SetFocus
    End If
End Sub
Private Sub cbo_operador_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_condicao.SetFocus
    End If
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If cbo_campo.ListIndex = -1 Then
        MsgBox "Informe o campo a ser testado.", vbInformation, "Atenção!"
        cbo_campo.SetFocus
    ElseIf cbo_operador.ListIndex = -1 Then
        MsgBox "Informe o operando a ser testado.", vbInformation, "Atenção!"
        cbo_operador.SetFocus
    ElseIf txt_condicao = "" Then
        MsgBox "Informe a condição testada.", vbInformation, "Atenção!"
        txt_condicao.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_ok_Click()
    Call GravaConfiguracaoConsulta(Me.name, cbo_campo.Text, cbo_operador.Text, txt_condicao.Text)
    AtualizaGrid
    MSFlexGrid1.SetFocus
End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub Form_Activate()
    If lFlagConsulta = 0 Then
        FormataGrid
        Inicializa
        txt_condicao.SetFocus
    Else
        lFlagConsulta = 0
    End If
End Sub
Private Sub Inicializa()
    Dim i As Integer
    Dim xCampo As String
    Dim xOperador As String
    
    xCampo = "Data"
    xOperador = "Maior Igual"
    txt_condicao.Text = Format(g_data_def, "dd/mm/yyyy")
    
    If BuscaConfiguracaoConsulta(Me.name) Then
        xCampo = RetiraGString(1)
        xOperador = RetiraGString(2)
        txt_condicao.Text = RetiraGString(3)
        g_string = ""
    End If
    
    cbo_campo.ListIndex = -1
    For i = 0 To cbo_campo.ListCount - 1
        cbo_campo.ListIndex = i
        If cbo_campo.Text = xCampo Then
            Exit For
        End If
    Next
    
    cbo_operador.ListIndex = -1
    For i = 0 To cbo_operador.ListCount - 1
        cbo_operador.ListIndex = i
        If cbo_operador.Text = xOperador Then
            Exit For
        End If
    Next
    AtualizaGrid
End Sub
Private Sub Form_Deactivate()
    lFlagConsulta = 1
End Sub
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
    CentraForm Me
    
    If g_nome_usuario = "L.M.C." Then
        lNomeTabela = "Entrada_Combustivel_LMC"
        Me.Caption = Me.Caption & " - LMC"
    Else
        lNomeTabela = "Entrada_Combustivel"
    End If
    PreencheCampos
    fConsultaPreencheOperador Me.cbo_operador
    g_string = ""
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub MSFLEXGRID1_DblClick()
    MarcaCelulas
    cmd_sair_Click
End Sub
Private Sub MSFLEXGRID1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        MarcaCelulas
        cmd_sair_Click
    ElseIf KeyCode = 98 Then
        KeyCode = 40
    ElseIf KeyCode = 99 Then
        KeyCode = 34
    ElseIf KeyCode = 104 Then
        KeyCode = 38
    ElseIf KeyCode = 105 Then
        KeyCode = 33
    End If
End Sub
Private Sub MSFLEXGRID1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
        KeyAscii = 0
        MarcaCelulas
        cmd_sair_Click
    End If
End Sub
Private Sub txt_condicao_GotFocus()
    txt_condicao.SelStart = 0
    txt_condicao.SelLength = 5
End Sub
Private Sub txt_condicao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok_Click
    End If
End Sub
