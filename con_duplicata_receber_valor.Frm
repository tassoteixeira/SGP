VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form consulta_duplicata_receber_valor 
   Caption         =   "Consulta de Duplicatas a Receber por Valor"
   ClientHeight    =   5355
   ClientLeft      =   1170
   ClientTop       =   2505
   ClientWidth     =   9270
   Icon            =   "con_duplicata_receber_valor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "con_duplicata_receber_valor.frx":0442
   ScaleHeight     =   5355
   ScaleWidth      =   9270
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   6780
      Picture         =   "con_duplicata_receber_valor.frx":0488
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Inicia a pesquisa selecionada."
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   8460
      Picture         =   "con_duplicata_receber_valor.frx":1A92
      Style           =   1  'Graphical
      TabIndex        =   6
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
      Width           =   3495
      Begin VB.ComboBox cbo_operador 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txt_condicao 
         Height          =   285
         Left            =   1920
         TabIndex        =   2
         Top             =   360
         Width           =   1395
      End
      Begin VB.Label Label4 
         Caption         =   "O&perador"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   780
         Width           =   1635
      End
      Begin VB.Label Label3 
         Caption         =   "&Valor do Vencimento"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   420
         Width           =   1635
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4035
      Left            =   60
      TabIndex        =   7
      Top             =   1260
      Width           =   9135
      _ExtentX        =   16113
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
Attribute VB_Name = "consulta_duplicata_receber_valor"
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
        lSQL = lSQL & "SELECT Duplicata_Receber.Empresa, Duplicata_Receber.[Numero do Documento], Duplicata_Receber.[Codigo do Cliente], Cliente.[Razao Social], Duplicata_Receber.[Data do Vencimento], Duplicata_Receber.[Valor do Vencimento], Duplicata_Receber.[Data de Emissao], Duplicata_Receber.[Data do Periodo Inicial], Duplicata_Receber.[Data do Periodo Final], Duplicata_Receber.[Numero da Duplicata], Duplicata_Receber.[Valor do Custo Bancario]"
        lSQL = lSQL & "  FROM Duplicata_Receber, Cliente"
        lSQL = lSQL & " WHERE Duplicata_Receber.Empresa = " & g_empresa
        lSQL = lSQL & "   AND Cliente.Codigo = Duplicata_Receber.[Codigo do Cliente]"
        'lSQL = lSQL & " And Duplicata_Receber.[Valor do Vencimento] " & x_operando & " " & x_condicao
        lSQL = lSQL & "   AND Duplicata_Receber." & fConsultaPreparaCondicao("Valor do Vencimento", 6, cbo_operador.Text, txt_condicao.Text)
        lSQL = lSQL & " ORDER BY Duplicata_receber.[Valor do Vencimento], Duplicata_receber.[Data do Vencimento]"
        'tempo = Time
        Set rsDados = Conectar.RsConexao(lSQL)
        FormataGrid
        PreencheGrid
        'MsgBox "Tempo gasto: " & DateDiff("s", tempo, Time)
    End If
    Exit Sub

ErroConsulta:
    If Err = 3075 Then
        MsgBox "Condição inválida.", vbInformation, "Erro de Consulta"
    Else
        MsgBox Error, vbInformation, "Erro de Consulta"
    End If
    Exit Sub




'    On Error GoTo ErroConsulta
'    Dim x_operando As String
'    Dim x_condicao As String
'    Dim x_data As Date
'    Dim x_nome_campo As String
'    x_nome_campo = "Valor do Vencimento"
'    'txt_condicao = Format(txt_condicao, "###,###,##0.00")
'    x_condicao = txt_condicao
'    If x_condicao = "" Then
'        x_condicao = 0
'    End If
'
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
'        l_campo = "Select Duplicata_Receber.Empresa, Duplicata_Receber.[Numero do Documento], Duplicata_Receber.[Codigo do Cliente], Cliente.[Razao Social], Duplicata_Receber.[Data do Vencimento], Duplicata_Receber.[Valor do Vencimento], Duplicata_Receber.[Data de Emissao], Duplicata_Receber.[Data do Periodo Inicial], Duplicata_Receber.[Data do Periodo Final], Duplicata_Receber.[Numero da Duplicata], Duplicata_Receber.[Valor do Custo Bancario]"
'        l_arquivo = " From " & tbl_tabela.name & ", Cliente"
'        l_condicao = " Where Cliente.Codigo = Duplicata_Receber.[Codigo do Cliente]"
'        l_condicao = l_condicao & " And Duplicata_Receber.Empresa <> 5"
'        l_condicao = l_condicao & " And Duplicata_Receber.[Valor do Vencimento] " & x_operando & " " & x_condicao
'        l_ordem = " order by Duplicata_receber.[Valor do Vencimento], Duplicata_receber.[Data do Vencimento]"
'        l_sql = l_campo & l_arquivo & l_condicao & l_ordem
'        dta_tabela.RecordSource = l_sql
'        dta_tabela.Refresh
'        FormaGrid
'    End If
'    Exit Sub
'ErroConsulta:
'    If Err = 3075 Then
'        MsgBox "Condição inválida.", 48, "Erro de Consulta"
'    Else
'        MsgBox Error, 48, "Erro de Consulta"
'    End If
'    Exit Sub
End Sub
Private Sub Finaliza()
    Set rsDados = Nothing
End Sub
'Private Sub FormaGrid()
'    Dim i As Integer
'    While TDBGrid1.Columns.Count <> 0
'        TDBGrid1.Columns.Remove 0
'    Wend
'    For i = 0 To 10
'        TDBGrid1.Columns.Add 0
'    Next
'    For i = 0 To 10
'        TDBGrid1.Columns(i).Visible = True
'    Next
'    TDBGrid1.Columns(0).DataField = "Empresa"
'    TDBGrid1.Columns(0).Caption = "Empresa"
'    TDBGrid1.Columns(0).Alignment = dbgLeft
'    TDBGrid1.Columns(0).HeadAlignment = dbgCenter
'    TDBGrid1.Columns(0).Width = 700
'    TDBGrid1.Columns(1).DataField = "Numero do Documento"
'    TDBGrid1.Columns(1).Caption = "N. Documento"
'    TDBGrid1.Columns(1).Alignment = dbgLeft
'    TDBGrid1.Columns(1).HeadAlignment = dbgCenter
'    TDBGrid1.Columns(1).Width = 900
'    TDBGrid1.Columns(2).DataField = "Codigo do Cliente"
'    TDBGrid1.Columns(2).Caption = "Código do Cliente"
'    TDBGrid1.Columns(2).Alignment = dbgRight
'    TDBGrid1.Columns(2).HeadAlignment = dbgCenter
'    TDBGrid1.Columns(2).Width = 800
'    TDBGrid1.Columns(3).DataField = "Razao Social"
'    TDBGrid1.Columns(3).Caption = "Razão Social"
'    TDBGrid1.Columns(3).Alignment = dbgLeft
'    TDBGrid1.Columns(3).HeadAlignment = dbgCenter
'    TDBGrid1.Columns(3).Width = 3000
'    TDBGrid1.Columns(4).DataField = "Data de Emissao"
'    TDBGrid1.Columns(4).NumberFormat = "General Date"
'    TDBGrid1.Columns(4).Caption = "Data de Emissão"
'    TDBGrid1.Columns(4).Alignment = dbgCenter
'    TDBGrid1.Columns(4).HeadAlignment = dbgCenter
'    TDBGrid1.Columns(4).Width = 1000
'    TDBGrid1.Columns(5).DataField = "Data do Vencimento"
'    TDBGrid1.Columns(5).NumberFormat = "General Date"
'    TDBGrid1.Columns(5).Caption = "Data do Vencimento"
'    TDBGrid1.Columns(5).Alignment = dbgCenter
'    TDBGrid1.Columns(5).HeadAlignment = dbgCenter
'    TDBGrid1.Columns(5).Width = 1000
'    TDBGrid1.Columns(6).DataField = "Valor do Vencimento"
'    TDBGrid1.Columns(6).Caption = "Valor do Vencimento"
'    TDBGrid1.Columns(6).Alignment = dbgRight
'    TDBGrid1.Columns(6).HeadAlignment = dbgCenter
'    TDBGrid1.Columns(6).Width = 1000
'    TDBGrid1.Columns(6).NumberFormat = "Currency"
'    TDBGrid1.Columns(7).DataField = "Data do Periodo Inicial"
'    TDBGrid1.Columns(7).NumberFormat = "General Date"
'    TDBGrid1.Columns(7).Caption = "Data Inicial"
'    TDBGrid1.Columns(7).Alignment = dbgCenter
'    TDBGrid1.Columns(7).HeadAlignment = dbgCenter
'    TDBGrid1.Columns(7).Width = 1000
'    TDBGrid1.Columns(8).DataField = "Data do Periodo Final"
'    TDBGrid1.Columns(8).NumberFormat = "General Date"
'    TDBGrid1.Columns(8).Caption = "Data Final"
'    TDBGrid1.Columns(8).Alignment = dbgCenter
'    TDBGrid1.Columns(8).HeadAlignment = dbgCenter
'    TDBGrid1.Columns(8).Width = 1000
'    TDBGrid1.Columns(9).DataField = "Numero da Duplicata"
'    TDBGrid1.Columns(9).Caption = "N. Duplicata"
'    TDBGrid1.Columns(9).Alignment = dbgLeft
'    TDBGrid1.Columns(9).HeadAlignment = dbgCenter
'    TDBGrid1.Columns(9).Width = 800
'    TDBGrid1.Columns(10).DataField = "Valor do Custo Bancario"
'    TDBGrid1.Columns(10).Caption = "Custo Bancário"
'    TDBGrid1.Columns(10).Alignment = dbgRight
'    TDBGrid1.Columns(10).HeadAlignment = dbgCenter
'    TDBGrid1.Columns(10).Width = 1000
'    TDBGrid1.Columns(10).NumberFormat = "Currency"
'End Sub
Private Sub FormataGrid()
    Dim i As Integer
    
    MSFlexGrid1.Visible = False
    MSFlexGrid1.WordWrap = True
    MSFlexGrid1.Cols = 11
    MSFlexGrid1.ColWidth(0) = 500
    MSFlexGrid1.Rows = 1
    i = 0
    MSFlexGrid1.TextMatrix(0, i) = "Empresa"
    MSFlexGrid1.ColAlignment(i) = flexAlignRightCenter
    MSFlexGrid1.ColWidth(i) = 700
    i = i + 1
    MSFlexGrid1.TextMatrix(0, i) = "N. Documento"
    MSFlexGrid1.ColAlignment(i) = flexAlignLeftCenter
    MSFlexGrid1.ColWidth(i) = 900
    i = i + 1
    MSFlexGrid1.TextMatrix(0, i) = "Código do Cliente"
    MSFlexGrid1.ColAlignment(i) = flexAlignRightCenter
    MSFlexGrid1.ColWidth(i) = 800
    i = i + 1
    MSFlexGrid1.TextMatrix(0, i) = "Razão Social"
    MSFlexGrid1.ColAlignment(i) = flexAlignLeftCenter
    MSFlexGrid1.ColWidth(i) = 3000
    i = i + 1
    MSFlexGrid1.TextMatrix(0, i) = "Data de Emissão"
    MSFlexGrid1.ColAlignment(i) = flexAlignCenterCenter
    MSFlexGrid1.ColWidth(i) = 1000
    i = i + 1
    MSFlexGrid1.TextMatrix(0, i) = "Data do Vencimento"
    MSFlexGrid1.ColAlignment(i) = flexAlignCenterCenter
    MSFlexGrid1.ColWidth(i) = 1000
    i = i + 1
    MSFlexGrid1.TextMatrix(0, i) = "Valor do Vencimento"
    MSFlexGrid1.ColAlignment(i) = flexAlignRightCenter
    MSFlexGrid1.ColWidth(i) = 1000
    i = i + 1
    MSFlexGrid1.TextMatrix(0, i) = "Data Inicial"
    MSFlexGrid1.ColAlignment(i) = flexAlignCenterCenter
    MSFlexGrid1.ColWidth(i) = 1000
    i = i + 1
    MSFlexGrid1.TextMatrix(0, i) = "Data Final"
    MSFlexGrid1.ColAlignment(i) = flexAlignCenterCenter
    MSFlexGrid1.ColWidth(i) = 1000
    i = i + 1
    MSFlexGrid1.TextMatrix(0, i) = "N. Duplicata"
    MSFlexGrid1.ColAlignment(i) = flexAlignLeftCenter
    MSFlexGrid1.ColWidth(i) = 1100
    i = i + 1
    MSFlexGrid1.TextMatrix(0, i) = "Custo Bancário"
    MSFlexGrid1.ColAlignment(i) = flexAlignRightCenter
    MSFlexGrid1.ColWidth(i) = 1000
End Sub
Private Sub MarcaCelulas()
    g_string = ""
    If MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 0) <> "" Then
        g_string = MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 0) & "|@|"
        g_string = g_string & MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 1) & "|@|"
        g_string = g_string & MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 2) & "|@|"
    End If
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
Private Sub cbo_operador_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_condicao.SetFocus
    End If
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If cbo_operador.ListIndex = -1 Then
        MsgBox "Informe o operando a ser testado.", 64, "Atenção!"
        cbo_operador.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_ok_Click()
    Call GravaConfiguracaoConsulta(Me.name, "Valor a Receber", cbo_operador.Text, txt_condicao.Text)
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
    Dim xCampo As String
    Dim xOperador As String
    
    xCampo = "Data do Vencimento"
    xOperador = "Maior Igual"
    txt_condicao.Text = Format(g_data_def, "dd/mm/yyyy")
    
    If BuscaConfiguracaoConsulta(Me.name) Then
        xCampo = RetiraGString(1)
        xOperador = RetiraGString(2)
        txt_condicao.Text = RetiraGString(3)
        g_string = ""
    End If
    
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
    
    lNomeTabela = "Duplicata_Receber"
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
    txt_condicao.SelLength = Len(txt_condicao.Text)
End Sub
Private Sub txt_condicao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Then
        KeyAscii = 46
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok_Click
    End If
End Sub
