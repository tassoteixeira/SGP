VERSION 5.00
Object = "{00028C21-0000-0000-0000-000000000046}#4.0#0"; "TDBG32.OCX"
Begin VB.Form consulta_movimento_saida_transferencia_produto 
   Caption         =   "Consulta Movimento de Saida de Transferência de Produtos"
   ClientHeight    =   5355
   ClientLeft      =   1455
   ClientTop       =   1785
   ClientWidth     =   8655
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "con_movimento_saida_transferencia_produto.frx":0000
   ScaleHeight     =   5355
   ScaleWidth      =   8655
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   7020
      Picture         =   "con_movimento_saida_transferencia_produto.frx":0046
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
      Picture         =   "con_movimento_saida_transferencia_produto.frx":1650
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   120
      Width           =   735
   End
   Begin VB.Data dta_tabela 
      Connect         =   "Access"
      DatabaseName    =   "Sgp_data.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Saida_Transferencia_Produto"
      Top             =   5040
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Frame frmSelecionar
      Height          =   1155
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6855
      Caption         =   "S&elecionar por..."
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
   Begin TrueDBGrid.TDBGrid TDBGrid1 
      Bindings        =   "con_movimento_saida_transferencia_produto.frx":2CE2
      Height          =   4035
      Left            =   60
      OleObjectBlob   =   "con_movimento_saida_transferencia_produto.frx":2CFB
      TabIndex        =   9
      Top             =   1260
      Width           =   8535
   End
End
Attribute VB_Name = "consulta_movimento_saida_transferencia_produto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tbl_tabela As Table
Dim tbl_produto As Table
Dim snp_campos As Snapshot
Dim l_campo As String
Dim l_arquivo As String
Dim l_condicao As String
Dim l_ordem As String
Dim l_sql As String
Private Sub AtualizaGrid()
    On Error GoTo ErroConsulta
    Dim x_operando As String
    Dim x_condicao As String
    Dim x_data As Date
    Dim x_nome_campo As String
    x_nome_campo = cbo_campo.Text
    If Mid(x_nome_campo, 1, 1) = "[" Then
        x_nome_campo = Mid(x_nome_campo, 2, Len(x_nome_campo) - 2)
    End If
    x_condicao = txt_condicao
    snp_campos.MoveFirst
    Do Until snp_campos.EOF
        If snp_campos!name = x_nome_campo Then
            If snp_campos!Type = 8 Then
                x_condicao = "#" & CDate(Format(x_condicao, "mm/dd/yyyy")) & "#"
            ElseIf snp_campos!Type = 10 And cbo_operador.Text = "Semelhante" Then
                x_condicao = Chr(34) & "*" & x_condicao & "*" & Chr(34)
            ElseIf snp_campos!Type = 10 And cbo_operador.Text <> "Semelhante" Then
                x_condicao = Chr(34) & x_condicao & Chr(34)
            End If
            Exit Do
        End If
        snp_campos.MoveNext
    Loop
    If cbo_operador.Text = "Diferente" Then
        x_operando = "<>"
    ElseIf cbo_operador.Text = "Igual" Then
        x_operando = "="
    ElseIf cbo_operador.Text = "Maior" Then
        x_operando = ">"
    ElseIf cbo_operador.Text = "Maior Igual" Then
        x_operando = ">="
    ElseIf cbo_operador.Text = "Menor" Then
        x_operando = "<"
    ElseIf cbo_operador.Text = "Menor Igual" Then
        x_operando = "<="
    ElseIf cbo_operador.Text = "Semelhante" Then
        x_operando = "Like"
    End If
    If ValidaCampos Then
        l_campo = "Select Saida_Transferencia_Produto.[Data da Transferencia], Saida_Transferencia_Produto.[Numero do Documento], Produto.Nome, Saida_Transferencia_Produto.[Preco de Custo], Saida_Transferencia_Produto.Quantidade, Empresas.Nome as NomeEmpresa, Saida_Transferencia_Produto.[Data da Digitacao], Saida_Transferencia_Produto.Observacao, Saida_Transferencia_Produto.[Codigo do Produto2]"
        l_arquivo = " From " & tbl_tabela.name & ", Produto, Empresas"
        If cbo_campo.ListIndex <> -1 Then
            l_condicao = ""
            l_condicao = " Where " & tbl_tabela.name & ".Empresa = " & g_empresa
            l_condicao = l_condicao & " And Produto.Codigo = Saida_Transferencia_Produto.[Codigo do Produto2]"
            l_condicao = l_condicao & " And Empresas.Codigo = Saida_Transferencia_Produto.[Entrou na Empresa]"
            l_condicao = l_condicao & " And " & tbl_tabela.name & "." & cbo_campo.Text & " " & x_operando & " " & x_condicao
        Else
            l_condicao = ""
        End If
        l_ordem = ""
        l_ordem = " order by [Data da Transferencia], [Codigo do Produto2], [Numero do Documento]"
        l_sql = l_campo & l_arquivo & l_condicao & l_ordem
        dta_tabela.RecordSource = l_sql
        dta_tabela.Refresh
        FormaGrid
    End If
    Exit Sub
ErroConsulta:
    If Err = 3075 Then
        MsgBox "Condição inválida.", 48, "Erro de Consulta"
    Else
        MsgBox Error, 48, "Erro de Consulta"
    End If
    Exit Sub
End Sub
Private Sub Finaliza()
    tbl_tabela.Close
    tbl_produto.Close
End Sub
Private Sub FormaGrid()
    Dim i As Integer
    While TDBGrid1.Columns.Count <> 0
        TDBGrid1.Columns.Remove 0
    Wend
    For i = 0 To 8
        TDBGrid1.Columns.Add 0
    Next
    For i = 0 To 7
        TDBGrid1.Columns(i).Visible = True
    Next
    TDBGrid1.Columns(0).DataField = "Data da Transferencia"
    TDBGrid1.Columns(0).NumberFormat = "General Date"
    TDBGrid1.Columns(0).Caption = "Data da Transferência"
    TDBGrid1.Columns(0).Alignment = dbgCenter
    TDBGrid1.Columns(0).HeadAlignment = dbgCenter
    TDBGrid1.Columns(0).Width = 1000
    TDBGrid1.Columns(1).DataField = "Numero do Documento"
    TDBGrid1.Columns(1).Caption = "Número do Documento"
    TDBGrid1.Columns(1).Alignment = dbgRight
    TDBGrid1.Columns(1).HeadAlignment = dbgCenter
    TDBGrid1.Columns(1).Width = 1000
    TDBGrid1.Columns(2).DataField = "Nome"
    TDBGrid1.Columns(2).Caption = "Produto"
    TDBGrid1.Columns(2).Alignment = dbgLeft
    TDBGrid1.Columns(2).HeadAlignment = dbgCenter
    TDBGrid1.Columns(2).Width = 2000
    TDBGrid1.Columns(3).DataField = "Preco de Custo"
    TDBGrid1.Columns(3).NumberFormat = "Currency"
    TDBGrid1.Columns(3).Caption = "Preço de Custo"
    TDBGrid1.Columns(3).Alignment = dbgRight
    TDBGrid1.Columns(3).HeadAlignment = dbgCenter
    TDBGrid1.Columns(3).Width = 700
    TDBGrid1.Columns(4).DataField = "Quantidade"
    TDBGrid1.Columns(4).NumberFormat = "Currency"
    TDBGrid1.Columns(4).Caption = "Quant."
    TDBGrid1.Columns(4).Alignment = dbgRight
    TDBGrid1.Columns(4).HeadAlignment = dbgCenter
    TDBGrid1.Columns(4).Width = 700
    TDBGrid1.Columns(5).DataField = "NomeEmpresa"
    TDBGrid1.Columns(5).Caption = "Entrou na Empresa"
    TDBGrid1.Columns(5).Alignment = dbgLeft
    TDBGrid1.Columns(5).HeadAlignment = dbgCenter
    TDBGrid1.Columns(5).Width = 2500
    TDBGrid1.Columns(6).DataField = "Data da Digitacao"
    TDBGrid1.Columns(6).NumberFormat = "General Date"
    TDBGrid1.Columns(6).Caption = "Data da Digitacao"
    TDBGrid1.Columns(6).Alignment = dbgCenter
    TDBGrid1.Columns(6).HeadAlignment = dbgCenter
    TDBGrid1.Columns(6).Width = 1000
    TDBGrid1.Columns(7).DataField = "Observacao"
    TDBGrid1.Columns(7).Caption = "Observação"
    TDBGrid1.Columns(7).Alignment = dbgLeft
    TDBGrid1.Columns(7).HeadAlignment = dbgCenter
    TDBGrid1.Columns(7).Width = 2000
    TDBGrid1.Columns(8).DataField = "Codigo do Produto2"
    TDBGrid1.Columns(8).Caption = "Código do Produto"
    TDBGrid1.Columns(8).Alignment = dbgRight
    TDBGrid1.Columns(8).HeadAlignment = dbgCenter
    TDBGrid1.Columns(8).Width = 700
End Sub
Private Sub MarcaCelulas()
    g_string = ""
    If TDBGrid1.Columns(0).Text <> "" Then
        g_string = TDBGrid1.Columns(0).Text & "|@|"
        g_string = g_string & TDBGrid1.Columns(8).Text & "|@|"
        g_string = g_string & TDBGrid1.Columns(1).Text & "|@|"
    End If
End Sub
Private Sub PreencheCampos()
    Dim i As Integer
    Dim x_string As String
    Set snp_campos = tbl_tabela.ListFields()
    cbo_campo.Clear
    snp_campos.MoveFirst
    Do Until snp_campos.EOF
        x_string = snp_campos!name
        For i = 1 To Len(snp_campos!name)
            If Mid(snp_campos!name, i, 1) = " " Then
                x_string = "[" & snp_campos!name & "]"
            End If
        Next
        cbo_campo.AddItem x_string
        snp_campos.MoveNext
    Loop
End Sub
Private Sub PreencheOperador()
    cbo_operador.Clear
    cbo_operador.AddItem "Diferente"
    cbo_operador.AddItem "Igual"
    cbo_operador.AddItem "Maior"
    cbo_operador.AddItem "Maior Igual"
    cbo_operador.AddItem "Menor"
    cbo_operador.AddItem "Menor Igual"
    cbo_operador.AddItem "Semelhante"
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
        MsgBox "Informe o campo a ser testado.", 64, "Atenção!"
        cbo_campo.SetFocus
    ElseIf cbo_operador.ListIndex = -1 Then
        MsgBox "Informe o operando a ser testado.", 64, "Atenção!"
        cbo_operador.SetFocus
    ElseIf txt_condicao = "" Then
        MsgBox "Informe a condição testada.", 64, "Atenção!"
        txt_condicao.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_ok_Click()
    AtualizaGrid
    TDBGrid1.SetFocus
End Sub
Private Sub cmd_sair_Click()
    Finaliza
    Unload Me
End Sub
Private Sub Form_Activate()
    TDBGrid1.Caption = tbl_tabela.name
    FormaGrid
    Inicializa
    txt_condicao.SetFocus
End Sub
Private Sub Inicializa()
    Dim i As Integer
    cbo_campo.ListIndex = -1
    For i = 0 To cbo_campo.ListCount - 1
        cbo_campo.ListIndex = i
        If cbo_campo = "[Data da Transferencia]" Then
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
    Set tbl_tabela = bd_sgp.OpenTable("Saida_Transferencia_Produto")
    Set tbl_produto = bd_sgp.OpenTable("Produto")
    PreencheCampos
    PreencheOperador
    g_string = ""
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
