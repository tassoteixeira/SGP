VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{00028C21-0000-0000-0000-000000000046}#4.0#0"; "TDBG32.OCX"
Begin VB.Form consulta_mala_direta 
   Caption         =   "Consulta de Mala Direta"
   ClientHeight    =   5355
   ClientLeft      =   2220
   ClientTop       =   2370
   ClientWidth     =   8655
   Icon            =   "con_mala_direta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "con_mala_direta.frx":0442
   ScaleHeight     =   5355
   ScaleWidth      =   8655
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   7020
      Picture         =   "con_mala_direta.frx":0488
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
      Picture         =   "con_mala_direta.frx":1A92
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   120
      Width           =   735
   End
   Begin VB.Data dta_tabela 
      Connect         =   "Access"
      DatabaseName    =   "\VB5\Sgp\Data\Sgp_data.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   60
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Mala_Direta"
      Top             =   4980
      Visible         =   0   'False
      Width           =   2055
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   1155
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6855
      _Version        =   65536
      _ExtentX        =   12091
      _ExtentY        =   2037
      _StockProps     =   14
      Caption         =   "S&elecionar por..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShadowStyle     =   1
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
      Bindings        =   "con_mala_direta.frx":3124
      Height          =   4035
      Left            =   60
      OleObjectBlob   =   "con_mala_direta.frx":313D
      TabIndex        =   9
      Top             =   1260
      Width           =   8535
   End
End
Attribute VB_Name = "consulta_mala_direta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim lFlagConsulta As Integer
    Dim tbl_tabela As Table
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
        If snp_campos!Name = x_nome_campo Then
            If snp_campos!Type = 8 Then
                x_condicao = "#" & Format(CDate(x_condicao), "yyyy/mm/dd") & "#"
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
        l_campo = "Select Codigo, Nome, Endereco, Bairro, Cidade, [Unidade Federativa], CEP, Telefone, [Data de Nascimento]"
        l_arquivo = " From " & tbl_tabela.Name
        If cbo_campo.ListIndex <> -1 Then
            l_condicao = ""
            l_condicao = " Where " & tbl_tabela.Name & "." & cbo_campo.Text & " " & x_operando & " " & x_condicao
        Else
            l_condicao = ""
        End If
        l_ordem = " order by Nome"
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
End Sub
Private Sub FormaGrid()
    'início das variáveis para traduzir
    'Dim VItem As New ValueItem
    'Dim VItems As ValueItems
    'fim das variáveis para traduzir
    Dim i As Integer
    While TDBGrid1.Columns.Count <> 0
        TDBGrid1.Columns.Remove 0
    Wend
    For i = 0 To 8
        TDBGrid1.Columns.Add 0
    Next
    For i = 0 To 8
        TDBGrid1.Columns(i).Visible = True
    Next
    i = 0
    TDBGrid1.Columns(i).DataField = "Codigo"
    TDBGrid1.Columns(i).Caption = "Código"
    TDBGrid1.Columns(i).Alignment = dbgRight
    TDBGrid1.Columns(i).HeadAlignment = dbgCenter
    TDBGrid1.Columns(i).Width = 550
    i = i + 1
    TDBGrid1.Columns(i).DataField = "Nome"
    TDBGrid1.Columns(i).Caption = "Nome"
    TDBGrid1.Columns(i).Alignment = dbgLeft
    TDBGrid1.Columns(i).HeadAlignment = dbgCenter
    TDBGrid1.Columns(i).Width = 3000
    i = i + 1
    TDBGrid1.Columns(i).DataField = "Endereco"
    TDBGrid1.Columns(i).Caption = "Endereço"
    TDBGrid1.Columns(i).Alignment = dbgLeft
    TDBGrid1.Columns(i).HeadAlignment = dbgCenter
    TDBGrid1.Columns(i).Width = 2000
    i = i + 1
    TDBGrid1.Columns(i).DataField = "Bairro"
    TDBGrid1.Columns(i).Caption = "Bairro"
    TDBGrid1.Columns(i).Alignment = dbgLeft
    TDBGrid1.Columns(i).HeadAlignment = dbgCenter
    TDBGrid1.Columns(i).Width = 2000
    i = i + 1
    TDBGrid1.Columns(i).DataField = "Cidade"
    TDBGrid1.Columns(i).Caption = "Cidade"
    TDBGrid1.Columns(i).Alignment = dbgLeft
    TDBGrid1.Columns(i).HeadAlignment = dbgCenter
    TDBGrid1.Columns(i).Width = 1600
    i = i + 1
    TDBGrid1.Columns(i).DataField = "Unidade Federativa"
    TDBGrid1.Columns(i).Caption = "UF"
    TDBGrid1.Columns(i).Alignment = dbgCenter
    TDBGrid1.Columns(i).HeadAlignment = dbgCenter
    TDBGrid1.Columns(i).Width = 300
    i = i + 1
    TDBGrid1.Columns(i).DataField = "CEP"
    TDBGrid1.Columns(i).Caption = "CEP"
    TDBGrid1.Columns(i).Alignment = dbgCenter
    TDBGrid1.Columns(i).HeadAlignment = dbgCenter
    TDBGrid1.Columns(i).Width = 800
    i = i + 1
    TDBGrid1.Columns(i).DataField = "Telefone"
    TDBGrid1.Columns(i).Caption = "Telefone"
    TDBGrid1.Columns(i).Alignment = dbgCenter
    TDBGrid1.Columns(i).HeadAlignment = dbgCenter
    TDBGrid1.Columns(i).Width = 1200
    i = i + 1
    TDBGrid1.Columns(i).DataField = "Data de Nascimento"
    TDBGrid1.Columns(i).Caption = "Data de Nascimento"
    TDBGrid1.Columns(i).NumberFormat = "General Date"
    TDBGrid1.Columns(i).Alignment = dbgCenter
    TDBGrid1.Columns(i).HeadAlignment = dbgCenter
    TDBGrid1.Columns(i).Width = 1200
    'Traduz Funcionário
    'Set VItems = TDBGrid1.Columns(14).ValueItems
    'With tbl_conta
    '    .MoveFirst
    '    Do Until .EOF
    '        VItem.Value = !Codigo
    '        VItem.DisplayValue = !Nome
    '        VItems.Add VItem
    '        .MoveNext
    '    Loop
    'End With
    'VItems.Translate = True
End Sub
Private Sub MarcaCelulas()
    g_string = ""
    If TDBGrid1.Columns(0).Text <> "" Then
        g_string = TDBGrid1.Columns(0).Text & "|@|"
    End If
End Sub
Private Sub PreencheCampos()
    Dim i As Integer
    Dim x_string As String
    Set snp_campos = tbl_tabela.ListFields()
    cbo_campo.Clear
    snp_campos.MoveFirst
    Do Until snp_campos.EOF
        x_string = snp_campos!Name
        For i = 1 To Len(snp_campos!Name)
            If Mid(snp_campos!Name, i, 1) = " " Then
                x_string = "[" & snp_campos!Name & "]"
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
    AtualizaGrid
    TDBGrid1.SetFocus
End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub Form_Activate()
    If lFlagConsulta = 0 Then
        TDBGrid1.Caption = tbl_tabela.Name
        FormaGrid
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
        If cbo_campo = "Nome" Then
            Exit For
        End If
    Next
    cbo_operador.ListIndex = -1
    For i = 0 To cbo_operador.ListCount - 1
        cbo_operador.ListIndex = i
        If cbo_operador = "Semelhante" Then
            Exit For
        End If
    Next
    txt_condicao = "A"
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
    Set tbl_tabela = bd_sgp.OpenTable("Mala_Direta")
    PreencheCampos
    PreencheOperador
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
    txt_condicao.SelLength = Len(txt_condicao)
End Sub
Private Sub txt_condicao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok_Click
    End If
End Sub
