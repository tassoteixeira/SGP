VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form consulta_medicao_combustiveis 
   Caption         =   "Consulta de Medição de Combustíveis"
   ClientHeight    =   5355
   ClientLeft      =   2220
   ClientTop       =   2370
   ClientWidth     =   8655
   Icon            =   "con_medicao_combustiveis.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "con_medicao_combustiveis.frx":0442
   ScaleHeight     =   5355
   ScaleWidth      =   8655
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
      Left            =   7020
      Picture         =   "con_medicao_combustiveis.frx":0488
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
      Picture         =   "con_medicao_combustiveis.frx":1A92
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
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   7117
      _Version        =   393216
      FixedCols       =   0
      BackColorSel    =   12582912
      ForeColorSel    =   16777215
      BackColorBkg    =   12632256
      SelectionMode   =   1
   End
End
Attribute VB_Name = "consulta_medicao_combustiveis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim lFlagConsulta As Integer
    Dim rsDados As New ADODB.Recordset
    Dim fld As ADODB.Field
    Dim lSQl As String
    Dim lNomeTabela As String
Private Sub AtualizaGrid()

    On Error GoTo ErroConsulta
    
    If ValidaCampos Then
        lSQl = ""
        lSQl = lSQl & "SELECT " & lNomeTabela & ".Data, " & lNomeTabela & ".[Numero do Tanque], " & lNomeTabela & ".[Tipo de Combustivel], " & lNomeTabela & ".Quantidade, " & lNomeTabela & ".[Observacao 1], " & lNomeTabela & ".[Observacao 2], " & lNomeTabela & ".[Observacao 3]"
        lSQl = lSQl & "  FROM " & lNomeTabela
        If cbo_campo.ListIndex <> -1 Then
            lSQl = lSQl & " WHERE " & lNomeTabela & ".Empresa = " & g_empresa
            lSQl = lSQl & "   AND " & lNomeTabela & "." & fConsultaPreparaCondicao(cbo_campo.Text, rsDados.Fields(cbo_campo.Text).Type, cbo_operador.Text, txt_condicao.Text)
        End If
        lSQl = lSQl & " ORDER BY Data, [Numero do Tanque]"
        Set rsDados = Conectar.RsConexao(lSQl)
        FormataGrid
        PreencheGrid
    End If
    Exit Sub
    
    
    
'    Dim x_operando As String
'    Dim x_condicao As String
'    Dim x_data As Date
'    Dim x_nome_campo As String
'    x_nome_campo = cbo_campo.Text
'    If Mid(x_nome_campo, 1, 1) = "[" Then
'        x_nome_campo = Mid(x_nome_campo, 2, Len(x_nome_campo) - 2)
'    End If
'    x_condicao = txt_condicao
'    snp_campos.MoveFirst
'    Do Until snp_campos.EOF
'        If snp_campos!Name = x_nome_campo Then
'            If snp_campos!Type = 8 Then
'                x_condicao = "#" & Format(CDate(x_condicao), "yyyy/mm/dd") & "#"
'            ElseIf snp_campos!Type = 10 And cbo_operador.Text = "Semelhante" Then
'                x_condicao = Chr(34) & "*" & x_condicao & "*" & Chr(34)
'            ElseIf snp_campos!Type = 10 And cbo_operador.Text <> "Semelhante" Then
'                x_condicao = Chr(34) & x_condicao & Chr(34)
'            End If
'            Exit Do
'        End If
'        snp_campos.MoveNext
'    Loop
'    If cbo_operador.Text = "Diferente" Then
'        x_operando = "<>"
'    ElseIf cbo_operador.Text = "Igual" Then
'        x_operando = "="
'    ElseIf cbo_operador.Text = "Maior" Then
'        x_operando = ">"
'    ElseIf cbo_operador.Text = "Maior Igual" Then
'        x_operando = ">="
'    ElseIf cbo_operador.Text = "Menor" Then
'        x_operando = "<"
'    ElseIf cbo_operador.Text = "Menor Igual" Then
'        x_operando = "<="
'    ElseIf cbo_operador.Text = "Semelhante" Then
'        x_operando = "Like"
'    End If
'    If ValidaCampos Then
'        l_campo = "SELECT Data, [Numero do Tanque], [Tipo de Combustivel], Quantidade, [Observacao 1], [Observacao 2], [Observacao 3]"
'        l_arquivo = " FROM " & tbl_tabela.Name
'        If cbo_campo.ListIndex <> -1 Then
'            l_condicao = ""
'            l_condicao = l_condicao & " WHERE " & tbl_tabela.Name & ".Empresa = " & g_empresa
'            l_condicao = l_condicao & " AND " & tbl_tabela.Name & "." & cbo_campo.Text & " " & x_operando & " " & x_condicao
'        Else
'            l_condicao = ""
'        End If
'        l_ordem = " ORDER BY Data, [Numero do Tanque]"
'        l_sql = l_campo & l_arquivo & l_condicao & l_ordem
'        dta_tabela.RecordSource = l_sql
'        dta_tabela.Refresh
'        FormataGrid
'    End If
'    Exit Sub

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
    MSFlexGrid1.Cols = 7
    MSFlexGrid1.ColWidth(0) = 500
    MSFlexGrid1.Rows = 1
    i = 0
    MSFlexGrid1.TextMatrix(0, i) = "Data"
    MSFlexGrid1.ColAlignment(i) = flexAlignCenterCenter
    MSFlexGrid1.ColWidth(i) = 1000
    i = i + 1
    MSFlexGrid1.TextMatrix(0, i) = "N.Tanque"
    MSFlexGrid1.ColAlignment(i) = flexAlignCenterCenter
    MSFlexGrid1.ColWidth(i) = 800
    i = i + 1
    MSFlexGrid1.TextMatrix(0, i) = "Tipo de Combustível"
    MSFlexGrid1.ColAlignment(i) = flexAlignCenterCenter
    MSFlexGrid1.ColWidth(i) = 1600
    i = i + 1
    MSFlexGrid1.TextMatrix(0, i) = "Quantidade"
    MSFlexGrid1.ColAlignment(i) = flexAlignRightCenter
    MSFlexGrid1.ColWidth(i) = 1000
    i = i + 1
    MSFlexGrid1.TextMatrix(0, i) = "Observação 1"
    MSFlexGrid1.ColAlignment(i) = flexAlignRightCenter
    MSFlexGrid1.ColWidth(i) = 1800
    i = i + 1
    MSFlexGrid1.TextMatrix(0, i) = "Observação 2"
    MSFlexGrid1.ColAlignment(i) = flexAlignRightCenter
    MSFlexGrid1.ColWidth(i) = 1800
    i = i + 1
    MSFlexGrid1.TextMatrix(0, i) = "Observação 3"
    MSFlexGrid1.ColAlignment(i) = flexAlignRightCenter
    MSFlexGrid1.ColWidth(i) = 1800
End Sub
Private Sub MarcaCelulas()
    g_string = ""
    If MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 0) <> "" Then
        g_string = MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 0) & "|@|"
        g_string = g_string & MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 1) & "|@|"
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
    lSQl = "SELECT * FROM " & lNomeTabela & " WHERE Empresa = 0"
    Set rsDados = Conectar.RsConexao(lSQl)
    cbo_campo.Clear
    For Each fld In rsDados.Fields
        cbo_campo.AddItem fld.Name
    Next
End Sub
Private Sub cbo_campo_GotFocus()
    SendMessageLong cbo_campo.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_campo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_operador.SetFocus
    End If
End Sub
Private Sub cbo_campo_LostFocus()
    SendMessageLong cbo_campo.hwnd, CB_SHOWDROPDOWN, False, 0
End Sub
Private Sub cbo_operador_GotFocus()
    SendMessageLong cbo_operador.hwnd, CB_SHOWDROPDOWN, True, 0
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
    Call GravaConfiguracaoConsulta(Me.Name, cbo_campo.Text, cbo_operador.Text, txt_condicao.Text)
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
    
    If BuscaConfiguracaoConsulta(Me.Name) Then
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
        lNomeTabela = "MedicaoCombustivelLMC"
        Me.Caption = Me.Caption & " - LMC"
    Else
        lNomeTabela = "MedicaoCombustivel"
    End If
    PreencheCampos
    fConsultaPreencheOperador Me.cbo_operador
    g_string = ""
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub MSFlexGrid1_DblClick()
    MarcaCelulas
    cmd_sair_Click
End Sub
Private Sub MSFlexGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
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
Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
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
