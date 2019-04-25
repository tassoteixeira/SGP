VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form consulta_movimento_vale_caixa 
   Caption         =   "Consulta Movimento de Vales de Caixa"
   ClientHeight    =   5355
   ClientLeft      =   1170
   ClientTop       =   2505
   ClientWidth     =   8655
   Icon            =   "con_movimento_vale_caixa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "con_movimento_vale_caixa.frx":0442
   ScaleHeight     =   5355
   ScaleWidth      =   8655
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   7020
      Picture         =   "con_movimento_vale_caixa.frx":0488
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
      Left            =   7860
      Picture         =   "con_movimento_vale_caixa.frx":1A92
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   120
      Width           =   735
   End
   Begin VB.Frame frmSelecionar 
      Caption         =   "S&elecionar por..."
      Height          =   1155
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6855
      Begin VB.ComboBox cbo_operador 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txt_condicao 
         Height          =   285
         Left            =   3120
         TabIndex        =   6
         Top             =   720
         Width           =   3615
      End
      Begin VB.ComboBox cbo_campo 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   300
         Width           =   5775
      End
      Begin VB.Label Label4 
         Caption         =   "O&perador"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   780
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Co&ndição"
         Height          =   255
         Left            =   2340
         TabIndex        =   5
         Top             =   780
         Width           =   795
      End
      Begin VB.Label Label2 
         Caption         =   "&Campo"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   675
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4035
      Left            =   60
      TabIndex        =   9
      Top             =   1260
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   7117
      _Version        =   393216
      FixedCols       =   0
      BackColorSel    =   12582912
      ForeColorSel    =   16777215
      BackColorBkg    =   12632256
      WordWrap        =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "consulta_movimento_vale_caixa"
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
        lSQL = lSQL & "SELECT Movimento_Vale_Caixa.Data, Movimento_Vale_Caixa.Periodo, Movimento_Vale_Caixa.[Numero da Ilha], Movimento_Vale_Caixa.[Tipo do Movimento], Movimento_Vale_Caixa.[Codigo do Funcionario], Funcionario.Nome as NomeFuncionario"
        lSQL = lSQL & "  FROM " & lNomeTabela & ", Funcionario"
        If cbo_campo.ListIndex <> -1 Then
            lSQL = lSQL & " WHERE " & lNomeTabela & ".Empresa = " & g_empresa
            lSQL = lSQL & "   AND " & lNomeTabela & ".Empresa = Funcionario.Empresa"
            lSQL = lSQL & "   AND " & lNomeTabela & ".[Codigo do Funcionario] = Funcionario.Codigo"
            If cbo_campo.Text = "Razao Social" Then
                lSQL = lSQL & "   AND Cliente." & fConsultaPreparaCondicao(cbo_campo.Text, 200, cbo_operador.Text, txt_condicao.Text)
            Else
                lSQL = lSQL & "   AND " & lNomeTabela & "." & fConsultaPreparaCondicao(cbo_campo.Text, rsDados.Fields(cbo_campo.Text).Type, cbo_operador.Text, txt_condicao.Text)
            End If
        End If
        lSQL = lSQL & " ORDER BY Cliente.[Razao Social], Duplicata_receber.[Data do Vencimento]"
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
'Private Sub FormaGrid()
'    'início das variáveis para traduzir
'    Dim VItem As New ValueItem
'    Dim VItems As ValueItems
'    'fim das variáveis para traduzir
'    Dim i As Integer
'    While TDBGrid1.Columns.Count <> 0
'        TDBGrid1.Columns.Remove 0
'    Wend
'    For i = 0 To 9
'        TDBGrid1.Columns.Add 0
'    Next
'    For i = 0 To 9
'        TDBGrid1.Columns(i).Visible = True
'    Next
'    TDBGrid1.Columns(0).DataField = "Data"
'    TDBGrid1.Columns(0).NumberFormat = "General Date"
'    TDBGrid1.Columns(0).Caption = "Data do Movimento"
'    TDBGrid1.Columns(0).Alignment = dbgCenter
'    TDBGrid1.Columns(0).HeadAlignment = dbgCenter
'    TDBGrid1.Columns(0).Width = 1000
'    TDBGrid1.Columns(1).DataField = "Tipo de Combustivel"
'    TDBGrid1.Columns(1).Caption = "Tipo de Combustível"
'    TDBGrid1.Columns(1).Alignment = dbgCenter
'    TDBGrid1.Columns(1).HeadAlignment = dbgCenter
'    TDBGrid1.Columns(1).Width = 1100
'    TDBGrid1.Columns(2).DataField = "Codigo do Fornecedor"
'    TDBGrid1.Columns(2).Caption = "Código do Fornecedor"
'    TDBGrid1.Columns(2).Alignment = dbgLeft
'    TDBGrid1.Columns(2).HeadAlignment = dbgCenter
'    TDBGrid1.Columns(2).Width = 1000
'    TDBGrid1.Columns(3).DataField = "nome"
'    TDBGrid1.Columns(3).Caption = "Nome do Fornecedor"
'    TDBGrid1.Columns(3).Alignment = dbgLeft
'    TDBGrid1.Columns(3).HeadAlignment = dbgCenter
'    TDBGrid1.Columns(3).Width = 2000
'    TDBGrid1.Columns(4).DataField = "Numero da Nota"
'    TDBGrid1.Columns(4).Caption = "Número da Nota"
'    TDBGrid1.Columns(4).Alignment = dbgLeft
'    TDBGrid1.Columns(4).HeadAlignment = dbgCenter
'    TDBGrid1.Columns(4).Width = 1000
'    TDBGrid1.Columns(5).DataField = "Valor do Litro"
'    TDBGrid1.Columns(5).Caption = "Valor do Litro"
'    TDBGrid1.Columns(5).Alignment = dbgRight
'    TDBGrid1.Columns(5).HeadAlignment = dbgCenter
'    TDBGrid1.Columns(5).Width = 800
''    TDBGrid1.Columns(5).NumberFormat = "Currency"
'    TDBGrid1.Columns(6).DataField = "Quantidade"
'    TDBGrid1.Columns(6).Caption = "Quantidade"
'    TDBGrid1.Columns(6).Alignment = dbgRight
'    TDBGrid1.Columns(6).HeadAlignment = dbgCenter
'    TDBGrid1.Columns(6).Width = 1000
'    TDBGrid1.Columns(6).NumberFormat = "Currency"
'    TDBGrid1.Columns(7).DataField = "Valor da Entrada"
'    TDBGrid1.Columns(7).Caption = "Valor da Entrada"
'    TDBGrid1.Columns(7).Alignment = dbgRight
'    TDBGrid1.Columns(7).HeadAlignment = dbgCenter
'    TDBGrid1.Columns(7).Width = 1000
'    TDBGrid1.Columns(7).NumberFormat = "Currency"
'    TDBGrid1.Columns(8).DataField = "Tipo de Transporte"
'    TDBGrid1.Columns(8).Caption = "Tipo de Transporte"
'    TDBGrid1.Columns(8).Alignment = dbgLeft
'    TDBGrid1.Columns(8).HeadAlignment = dbgCenter
'    TDBGrid1.Columns(8).Width = 1000
'    TDBGrid1.Columns(9).DataField = "Numero do Tanque"
'    TDBGrid1.Columns(9).Caption = "Número do Tanque"
'    TDBGrid1.Columns(9).Alignment = dbgCenter
'    TDBGrid1.Columns(9).HeadAlignment = dbgCenter
'    TDBGrid1.Columns(9).Width = 1000
'    'Início da Tradução
'    Set VItems = TDBGrid1.Columns(8).ValueItems
'    VItem.Value = 0
'    VItem.DisplayValue = "Entrega"
'    VItems.Add VItem
'    VItem.Value = 1
'    VItem.DisplayValue = "Retirada"
'    VItems.Add VItem
'    VItems.Translate = True
'End Sub
Private Sub FormataGrid()
    Dim i As Integer
    
    MSFlexGrid1.Visible = False
    MSFlexGrid1.WordWrap = True
    MSFlexGrid1.Cols = 10
    MSFlexGrid1.ColWidth(0) = 500
    MSFlexGrid1.Rows = 1
    i = 0
    MSFlexGrid1.TextMatrix(0, i) = "Data"
    MSFlexGrid1.ColAlignment(i) = flexAlignLeftCenter
    MSFlexGrid1.ColWidth(i) = 900
    i = i + 1
    MSFlexGrid1.TextMatrix(0, i) = "Período"
    MSFlexGrid1.ColAlignment(i) = flexAlignCenterCenter
    MSFlexGrid1.ColWidth(i) = 500
    i = i + 1
    MSFlexGrid1.TextMatrix(0, i) = "Ilha"
    MSFlexGrid1.ColAlignment(i) = flexAlignCenterCenter
    MSFlexGrid1.ColWidth(i) = 500
    i = i + 1
    MSFlexGrid1.TextMatrix(0, i) = "Tipo do Movimento"
    MSFlexGrid1.ColAlignment(i) = flexAlignCenterCenter
    MSFlexGrid1.ColWidth(i) = 500
    i = i + 1
    MSFlexGrid1.TextMatrix(0, i) = "Código do Funcionário"
    MSFlexGrid1.ColAlignment(i) = flexAlignRightCenter
    MSFlexGrid1.ColWidth(i) = 800
    i = i + 1
    MSFlexGrid1.TextMatrix(0, i) = "Nome"
    MSFlexGrid1.ColAlignment(i) = flexAlignLeftCenter
    MSFlexGrid1.ColWidth(i) = 3000
End Sub
Private Sub MarcaCelulas()
    g_string = ""
    If MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 0) <> "" Then
        g_string = MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 0) & "|@|"
        g_string = g_string & MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 1) & "|@|"
        g_string = g_string & MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 4) & "|@|"
    End If
End Sub
Private Sub PreencheCampos()
    lSQL = "SELECT * FROM " & lNomeTabela & " WHERE Empresa = 0"
    Set rsDados = Conectar.RsConexao(lSQL)
    cbo_campo.Clear
    For Each fld In rsDados.Fields
        cbo_campo.AddItem fld.name
    Next
    cbo_campo.AddItem "Razao Social"
End Sub
Private Sub PreencheGrid()
    Dim i As Integer
    Dim i2 As Integer
    
    On Error GoTo ErroRotina
    
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
    MSFlexGrid1.Visible = True
    Exit Sub

ErroRotina:
    If Err = 3075 Then
        MsgBox "Condição inválida.", vbInformation, "Erro de Consulta"
    ElseIf Err = 3704 Then
        MsgBox "Condição incompatível.", vbInformation, "Erro de Consulta!"
    Else
        MsgBox Error, vbInformation, "Erro de Consulta"
    End If
    Exit Sub
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
    If MSFlexGrid1.Visible Then
        MSFlexGrid1.SetFocus
    End If
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
    cbo_campo.ListIndex = -1
    For i = 0 To cbo_campo.ListCount - 1
        cbo_campo.ListIndex = i
        If cbo_campo = "Data" Then
            Exit For
        End If
    Next
    cbo_operador.ListIndex = -1
    For i = 0 To cbo_operador.ListCount - 1
        cbo_operador.ListIndex = i
        If cbo_operador = "Maior Igual" Then
            Exit For
        End If
    Next
    txt_condicao = Format(g_data_def, "dd/mm/yyyy")
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
    
    lNomeTabela = "Movimento_Vale_Caixa"
    PreencheCampos
    fConsultaPreencheOperador Me.cbo_operador
    g_string = ""
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub TDBGrid1_DblClick()
    MarcaCelulas
    cmd_sair_Click
End Sub
Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
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
Private Sub TDBGrid1_KeyPress(KeyAscii As Integer)
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
