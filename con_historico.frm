VERSION 5.00
Object = "{00028C21-0000-0000-0000-000000000046}#4.0#0"; "TDBG32.OCX"
Begin VB.Form consulta_historico 
   Caption         =   "Consulta de Histórico"
   ClientHeight    =   5355
   ClientLeft      =   1245
   ClientTop       =   2385
   ClientWidth     =   8655
   Icon            =   "con_historico.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "con_historico.frx":0442
   ScaleHeight     =   5355
   ScaleWidth      =   8655
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   7020
      Picture         =   "con_historico.frx":0488
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
      Picture         =   "con_historico.frx":1A92
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
      Left            =   60
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Movimento_Historico"
      Top             =   4980
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Frame Frame1 
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
   Begin TrueDBGrid.TDBGrid TDBGrid1 
      Bindings        =   "con_historico.frx":3124
      Height          =   4035
      Left            =   60
      OleObjectBlob   =   "con_historico.frx":313D
      TabIndex        =   9
      Top             =   1260
      Width           =   8535
   End
End
Attribute VB_Name = "consulta_historico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim flag_consulta_cheque As Integer
    Dim tbl_tabela As Table
    Dim tbl_funcionario As Table
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
        l_campo = "Select Movimento_Historico.Data, Movimento_Historico.Periodo, Movimento_Historico.[Numero da Ilha], Movimento_Historico.[Tipo do Movimento], Movimento_Historico.[Cheque Pre-Datado], Movimento_Historico.[Cheque A Vista], Movimento_Historico.Dinheiro, Movimento_Historico.Nota, Movimento_Historico.Amex, Movimento_Historico.Visa, Movimento_Historico.Dinners, Movimento_Historico.Hipercheque, Movimento_Historico.Assalto, Movimento_Historico.Afericao, Movimento_Historico.Transferencia, Movimento_Historico.Total, Movimento_Historico.[Codigo do Funcionario], Funcionario.Nome"
        l_arquivo = " From " & tbl_tabela.name & ", Funcionario"
        If cbo_campo.ListIndex <> -1 Then
            l_condicao = ""
            l_condicao = " Where " & tbl_tabela.name & ".Empresa = " & g_empresa
            l_condicao = l_condicao & " And Funcionario.Empresa = Movimento_Historico.Empresa"
            l_condicao = l_condicao & " And Funcionario.Codigo = Movimento_Historico.[Codigo do Funcionario]"
            l_condicao = l_condicao & " And " & tbl_tabela.name & "." & cbo_campo.Text & " " & x_operando & " " & x_condicao
        Else
            l_condicao = ""
        End If
        l_ordem = " order by Data, Movimento_Historico.Periodo, Movimento_Historico.[Numero da Ilha], [Tipo do Movimento]"
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
    tbl_funcionario.Close
End Sub
Private Sub FormaGrid()
    Dim i As Integer
    While TDBGrid1.Columns.Count <> 0
        TDBGrid1.Columns.Remove 0
    Wend
    For i = 0 To 17
        TDBGrid1.Columns.Add 0
    Next
    For i = 0 To 17
        TDBGrid1.Columns(i).Visible = True
    Next
    TDBGrid1.Columns(0).DataField = "Data"
    TDBGrid1.Columns(0).NumberFormat = "General Date"
    TDBGrid1.Columns(0).Caption = "Data"
    TDBGrid1.Columns(0).Alignment = dbgCenter
    TDBGrid1.Columns(0).HeadAlignment = dbgCenter
    TDBGrid1.Columns(0).Width = 1000
    TDBGrid1.Columns(1).DataField = "Periodo"
    TDBGrid1.Columns(1).Caption = "Período"
    TDBGrid1.Columns(1).Alignment = dbgCenter
    TDBGrid1.Columns(1).HeadAlignment = dbgCenter
    TDBGrid1.Columns(1).Width = 800
    TDBGrid1.Columns(2).DataField = "Numero da Ilha"
    TDBGrid1.Columns(2).Caption = "Número da Ilha"
    TDBGrid1.Columns(2).Alignment = dbgCenter
    TDBGrid1.Columns(2).HeadAlignment = dbgCenter
    TDBGrid1.Columns(2).Width = 800
    TDBGrid1.Columns(3).DataField = "Tipo do Movimento"
    TDBGrid1.Columns(3).Caption = "Tipo do Movimento"
    TDBGrid1.Columns(3).Alignment = dbgCenter
    TDBGrid1.Columns(3).HeadAlignment = dbgCenter
    TDBGrid1.Columns(3).Width = 950
    TDBGrid1.Columns(4).DataField = "Cheque Pre-Datado"
    TDBGrid1.Columns(4).Caption = "Cheque Pré-Datado"
    TDBGrid1.Columns(4).Alignment = dbgRight
    TDBGrid1.Columns(4).HeadAlignment = dbgCenter
    TDBGrid1.Columns(4).Width = 1000
    TDBGrid1.Columns(4).NumberFormat = "Currency"
    TDBGrid1.Columns(5).DataField = "Cheque A Vista"
    TDBGrid1.Columns(5).Caption = "Cheque à Vista"
    TDBGrid1.Columns(5).Alignment = dbgRight
    TDBGrid1.Columns(5).HeadAlignment = dbgCenter
    TDBGrid1.Columns(5).Width = 1000
    TDBGrid1.Columns(5).NumberFormat = "Currency"
    TDBGrid1.Columns(6).DataField = "Dinheiro"
    TDBGrid1.Columns(6).Caption = "Valor em Dinheiro"
    TDBGrid1.Columns(6).Alignment = dbgRight
    TDBGrid1.Columns(6).HeadAlignment = dbgCenter
    TDBGrid1.Columns(6).Width = 1000
    TDBGrid1.Columns(6).NumberFormat = "Currency"
    TDBGrid1.Columns(7).DataField = "Nota"
    TDBGrid1.Columns(7).Caption = "Valor em Nota"
    TDBGrid1.Columns(7).Alignment = dbgRight
    TDBGrid1.Columns(7).HeadAlignment = dbgCenter
    TDBGrid1.Columns(7).Width = 1000
    TDBGrid1.Columns(7).NumberFormat = "Currency"
    TDBGrid1.Columns(8).DataField = "Amex"
    TDBGrid1.Columns(8).Caption = "Valor em Amex/Sollo"
    TDBGrid1.Columns(8).Alignment = dbgRight
    TDBGrid1.Columns(8).HeadAlignment = dbgCenter
    TDBGrid1.Columns(8).Width = 1000
    TDBGrid1.Columns(8).NumberFormat = "Currency"
    TDBGrid1.Columns(9).DataField = "Visa"
    TDBGrid1.Columns(9).Caption = "Valor em Visa"
    TDBGrid1.Columns(9).Alignment = dbgRight
    TDBGrid1.Columns(9).HeadAlignment = dbgCenter
    TDBGrid1.Columns(9).Width = 1000
    TDBGrid1.Columns(9).NumberFormat = "Currency"
    TDBGrid1.Columns(10).DataField = "Dinners"
    TDBGrid1.Columns(10).Caption = "Vlr.CredCard /Dinners"
    TDBGrid1.Columns(10).Alignment = dbgRight
    TDBGrid1.Columns(10).HeadAlignment = dbgCenter
    TDBGrid1.Columns(10).Width = 1000
    TDBGrid1.Columns(10).NumberFormat = "Currency"
    TDBGrid1.Columns(11).DataField = "Hipercheque"
    TDBGrid1.Columns(11).Caption = "Valor em Hipercheque"
    TDBGrid1.Columns(11).Alignment = dbgRight
    TDBGrid1.Columns(11).HeadAlignment = dbgCenter
    TDBGrid1.Columns(11).Width = 1000
    TDBGrid1.Columns(11).NumberFormat = "Currency"
    TDBGrid1.Columns(12).DataField = "Assalto"
    TDBGrid1.Columns(12).Caption = "Valor em Assalto"
    TDBGrid1.Columns(12).Alignment = dbgRight
    TDBGrid1.Columns(12).HeadAlignment = dbgCenter
    TDBGrid1.Columns(12).Width = 1000
    TDBGrid1.Columns(12).NumberFormat = "Currency"
    TDBGrid1.Columns(13).DataField = "Afericao"
    TDBGrid1.Columns(13).Caption = "Valor em Aferição"
    TDBGrid1.Columns(13).Alignment = dbgRight
    TDBGrid1.Columns(13).HeadAlignment = dbgCenter
    TDBGrid1.Columns(13).Width = 1000
    TDBGrid1.Columns(13).NumberFormat = "Currency"
    TDBGrid1.Columns(14).DataField = "Transferencia"
    TDBGrid1.Columns(14).Caption = "Valor em Transferência"
    TDBGrid1.Columns(14).Alignment = dbgRight
    TDBGrid1.Columns(14).HeadAlignment = dbgCenter
    TDBGrid1.Columns(14).Width = 1100
    TDBGrid1.Columns(14).NumberFormat = "Currency"
    TDBGrid1.Columns(15).DataField = "Total"
    TDBGrid1.Columns(15).Caption = "Total"
    TDBGrid1.Columns(15).Alignment = dbgRight
    TDBGrid1.Columns(15).HeadAlignment = dbgCenter
    TDBGrid1.Columns(15).Width = 1000
    TDBGrid1.Columns(15).NumberFormat = "Currency"
    TDBGrid1.Columns(16).DataField = "Codigo do Funcionario"
    TDBGrid1.Columns(16).Caption = "Cod."
    TDBGrid1.Columns(16).Alignment = dbgRight
    TDBGrid1.Columns(16).HeadAlignment = dbgCenter
    TDBGrid1.Columns(16).Width = 800
    TDBGrid1.Columns(17).DataField = "Nome"
    TDBGrid1.Columns(17).Caption = "Nome"
    TDBGrid1.Columns(17).Alignment = dbgLeft
    TDBGrid1.Columns(17).HeadAlignment = dbgCenter
    TDBGrid1.Columns(17).Width = 3000
End Sub
Private Sub MarcaCelulas()
    g_string = ""
    If TDBGrid1.Columns(0).Text <> "" Then
        g_string = TDBGrid1.Columns(0).Text & "|@|"
        g_string = g_string & TDBGrid1.Columns(1).Text & "|@|"
        g_string = g_string & TDBGrid1.Columns(2).Text & "|@|"
        g_string = g_string & TDBGrid1.Columns(3).Text & "|@|"
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
    If flag_consulta_cheque = 0 Then
        TDBGrid1.Caption = tbl_tabela.name
        FormaGrid
        Inicializa
        txt_condicao.SetFocus
    Else
        flag_consulta_cheque = 0
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
    flag_consulta_cheque = 1
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
    Set tbl_tabela = bd_sgp.OpenTable("Movimento_Historico")
    Set tbl_funcionario = bd_sgp.OpenTable("Funcionario")
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
    txt_condicao.SelLength = 5
End Sub
Private Sub txt_condicao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok_Click
    End If
End Sub
