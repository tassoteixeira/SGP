VERSION 5.00
Object = "{00028C21-0000-0000-0000-000000000046}#4.0#0"; "TDBG32.OCX"
Begin VB.Form super_consulta 
   Caption         =   "Super Consulta"
   ClientHeight    =   7800
   ClientLeft      =   2310
   ClientTop       =   885
   ClientWidth     =   7815
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "super_consulta.frx":0000
   ScaleHeight     =   7800
   ScaleWidth      =   7815
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   7020
      Picture         =   "super_consulta.frx":0046
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Inicia a pesquisa selecionada."
      Top             =   120
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   7020
      Picture         =   "super_consulta.frx":1320
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   1020
      Width           =   795
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
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   7500
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Frame frmTabela
      Height          =   435
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   6855
      Begin VB.ComboBox cbo_tabela 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   60
         Width           =   5775
      End
      Begin VB.Label Label1 
         Caption         =   "&Tabelas"
         Height          =   255
         Left            =   60
         TabIndex        =   1
         Top             =   120
         Width           =   675
      End
   End
   Begin VB.Frame frmSelecionar 
      Caption         =   "&Selecionar por..."
      Height          =   1155
      Left            =   60
      TabIndex        =   3
      Top             =   720
      Width           =   6855
      Begin VB.ComboBox cbo_operador 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txt_condicao 
         Height          =   285
         Left            =   3120
         TabIndex        =   9
         Top             =   720
         Width           =   3615
      End
      Begin VB.ComboBox cbo_campo 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   300
         Width           =   5775
      End
      Begin VB.Label Label4 
         Caption         =   "O&perador"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   780
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Co&ndição"
         Height          =   255
         Left            =   2340
         TabIndex        =   8
         Top             =   780
         Width           =   795
      End
      Begin VB.Label Label2 
         Caption         =   "&Campo"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   675
      End
   End
   Begin TrueDBGrid.TDBGrid TDBGrid1 
      Bindings        =   "super_consulta.frx":25FA
      Height          =   5895
      Left            =   0
      OleObjectBlob   =   "super_consulta.frx":2613
      TabIndex        =   10
      Top             =   1920
      Width           =   7815
   End
End
Attribute VB_Name = "super_consulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim snp_tabelas As Snapshot
    Dim snp_campos As Snapshot
    Dim l_nome_tabela As Table
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
    x_condicao = txt_condicao
    snp_campos.MoveFirst
    Do Until snp_campos.EOF
        If Mid(cbo_campo.Text, 1, 1) = "[" And snp_campos!name = Mid(cbo_campo.Text, 2, Len(cbo_campo.Text) - 2) Or snp_campos!name = cbo_campo.Text Then
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
        l_campo = "Select " & "* "
        l_arquivo = "From " & cbo_tabela.Text & " "
        If cbo_campo.ListIndex <> -1 Then
            l_condicao = "Where "
            l_condicao = l_condicao & cbo_tabela.Text & "." & cbo_campo.Text & " " & x_operando & " " & x_condicao
        Else
            l_condicao = ""
        End If
        'l_ordem = "order by cheques.emitente"
        l_ordem = ""
        l_sql = l_campo & l_arquivo & l_condicao & l_ordem
'TESTE        Set l_nome_tabela = bd_sgp.CreateSnapshot(l_sql)
        dta_tabela.RecordSource = l_sql
        dta_tabela.Refresh
    End If
    Exit Sub
ErroConsulta:
    If Err = 3075 Then
        MsgBox "Condição inválida.", 48, "Erro de Consulta"
        Exit Sub
    End If
    Exit Sub
End Sub
Private Sub Finaliza()
    tbl_cheque.Close
End Sub
Private Sub PreencheCampos()
    Set snp_campos = l_nome_tabela.ListFields()
    Dim i As Integer
    Dim x_string As String
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
Private Sub cbo_operador_GotFocus()
    SendMessageLong cbo_operador.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_operador_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_condicao.SetFocus
    End If
End Sub
Private Sub cbo_tabela_Click()
    If cbo_tabela.ListIndex <> -1 Then
        Set l_nome_tabela = bd_sgp.OpenTable(cbo_tabela.Text)
        dta_tabela.RecordSource = cbo_tabela.Text
        TDBGrid1.Caption = cbo_tabela.Text
        PreencheCampos
    End If
End Sub
Private Sub cbo_tabela_GotFocus()
    SendMessageLong cbo_tabela.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_tabela_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_campo.SetFocus
    End If
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If cbo_tabela.ListIndex = -1 Then
        MsgBox "Selecione a tabela.", 64, "Atenção!"
        cbo_tabela.SetFocus
    ElseIf cbo_campo.ListIndex = -1 Then
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
    Set snp_tabelas = bd_sgp.ListTables()
    cbo_tabela.Clear
    snp_tabelas.MoveFirst
    Do Until snp_tabelas.EOF
        If Mid(snp_tabelas!name, 1, 4) <> "MSys" Then
            cbo_tabela.AddItem snp_tabelas!name
        End If
        snp_tabelas.MoveNext
    Loop
    PreencheOperador
End Sub
Private Sub Form_Load()
    Screen.MousePointer = 1
    CentraForm Me
    Set tbl_cheque = bd_sgp.OpenTable("Cliente")
End Sub
Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 98 Then
        KeyCode = 40
    ElseIf KeyCode = 99 Then
        KeyCode = 34
    ElseIf KeyCode = 104 Then
        KeyCode = 38
    ElseIf KeyCode = 105 Then
        KeyCode = 33
    End If
End Sub
Private Sub txt_condicao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok_Click
    End If
End Sub
