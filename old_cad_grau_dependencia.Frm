VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{00028C21-0000-0000-0000-000000000046}#4.0#0"; "TDBG32.OCX"
Begin VB.Form cadastro_grau_dependencia 
   Caption         =   "Cadastro de Grau de Dependência"
   ClientHeight    =   4695
   ClientLeft      =   1095
   ClientTop       =   1455
   ClientWidth     =   6135
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "cad_grau_dependencia.frx":0000
   ScaleHeight     =   4695
   ScaleWidth      =   6135
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   120
      Picture         =   "cad_grau_dependencia.frx":0446
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Cria um novo registro."
      Top             =   3720
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   1020
      Picture         =   "cad_grau_dependencia.frx":1720
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Altera o registro atual."
      Top             =   3720
      Width           =   795
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   1920
      Picture         =   "cad_grau_dependencia.frx":29FA
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Exclui o registro atual."
      Top             =   3720
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   2820
      Picture         =   "cad_grau_dependencia.frx":3CD4
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   3720
      Width           =   795
   End
   Begin VB.Data dta_grau_dependencia 
      Caption         =   "dta_grau_dependencia"
      Connect         =   "Access"
      DatabaseName    =   "Sgp_data.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Grau_Dependencia"
      Top             =   3300
      Visible         =   0   'False
      Width           =   2115
   End
   Begin Threed.SSFrame frm_dados 
      Height          =   1155
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   5895
      _Version        =   65536
      _ExtentX        =   10398
      _ExtentY        =   2037
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      Begin VB.TextBox txt_codigo 
         Height          =   285
         Left            =   1740
         MaxLength       =   2
         TabIndex        =   2
         Top             =   300
         Width           =   375
      End
      Begin VB.TextBox txt_nome 
         Height          =   285
         Left            =   1740
         MaxLength       =   30
         TabIndex        =   4
         Top             =   660
         Width           =   3975
      End
      Begin VB.Label Label1 
         Caption         =   "Código"
         Height          =   255
         Left            =   180
         TabIndex        =   1
         Top             =   300
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Nome"
         Height          =   255
         Left            =   180
         TabIndex        =   3
         Top             =   660
         Width           =   1455
      End
   End
   Begin TrueDBGrid.TDBGrid tdbgrid_grau_dependencia 
      Bindings        =   "cad_grau_dependencia.frx":4FAE
      Height          =   2295
      Left            =   120
      OleObjectBlob   =   "cad_grau_dependencia.frx":4FCD
      TabIndex        =   5
      Top             =   1320
      Width           =   5895
   End
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   3840
      TabIndex        =   12
      Top             =   3600
      Width           =   2175
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "cad_grau_dependencia.frx":632D
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "cad_grau_dependencia.frx":7607
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "cad_grau_dependencia.frx":88E1
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "cad_grau_dependencia.frx":9BBB
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   5220
      Picture         =   "cad_grau_dependencia.frx":AE95
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Cancela o registro atual."
      Top             =   3720
      Width           =   795
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   4320
      Picture         =   "cad_grau_dependencia.frx":C16F
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Confirma o registro atual."
      Top             =   3720
      Width           =   795
   End
End
Attribute VB_Name = "cadastro_grau_dependencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flag_cadastro_grau_dependencia As Integer
Dim l_opcao As Integer
Dim l_codigo As Integer
Private Sub AtualTabe()
    l_codigo = Val(txt_codigo)
    With dta_grau_dependencia.Recordset
        !Codigo = Val(txt_codigo)
        !Nome = txt_nome
    End With
End Sub
Private Sub AtualTela()
    With dta_grau_dependencia.Recordset
        l_codigo = !Codigo
        txt_codigo = Format(!Codigo, "00")
        txt_nome = !Nome
    End With
    frm_dados.Enabled = False
End Sub
Private Sub Finaliza()
    dta_grau_dependencia.Recordset.Close
End Sub
Private Sub Inclui()
    l_opcao = 1
    DesativaBotoes
    cmd_novo.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    txt_codigo = 0
    If dta_grau_dependencia.Recordset.RecordCount > 0 Then
        dta_grau_dependencia.Recordset.Seek "<", 9999
        txt_codigo = dta_grau_dependencia.Recordset!Codigo
    End If
    txt_codigo = Format(Val(txt_codigo) + 1, "00")
End Sub
Private Sub cmd_alterar_Click()
    l_opcao = 2
    DesativaBotoes
    cmd_alterar.Visible = True
    cmd_alterar.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    frm_dados.Enabled = True
    txt_nome.SetFocus
End Sub
Private Sub cmd_anterior_Click()
    If dta_grau_dependencia.Recordset.RecordCount > 0 Then
        dta_grau_dependencia.Recordset.MovePrevious
        If dta_grau_dependencia.Recordset.BOF Then
            MsgBox "Início de Arquivo.", 48, "Atenção!"
            dta_grau_dependencia.Recordset.MoveFirst
            cmd_proximo.SetFocus
        End If
        AtualTela
    End If
End Sub
Private Sub cmd_cancelar_Click()
    LimpaTela
    l_opcao = 0
    If BuscaRegistro(l_codigo) Then
        AtivaBotoes
        cmd_alterar.SetFocus
    Else
        DesativaBotoes
        cmd_novo.Enabled = True
        cmd_sair.Enabled = True
        cmd_novo.SetFocus
    End If
End Sub
Private Sub LimpaTela()
    txt_codigo = ""
    txt_nome = ""
End Sub
Private Sub cmd_excluir_Click()
    If dta_grau_dependencia.Recordset!Codigo > 0 Then
        If (MsgBox("Deseja excluir este registro?", 4 + 32 + 256, "Atenção!")) = 6 Then
            dta_grau_dependencia.Recordset.Edit
            dta_grau_dependencia.Recordset.Delete
            LimpaTela
            If Not BuscaDados Then
                DesativaBotoes
                cmd_novo.Enabled = True
                cmd_sair.Enabled = True
                cmd_novo.SetFocus
            End If
        End If
    End If
End Sub
Private Sub cmd_novo_Click()
    LimpaTela
    Inclui
    frm_dados.Enabled = True
    txt_nome.SetFocus
End Sub
Private Sub cmd_ok_Click()
    On Error GoTo FileError
    If ValidaCampos Then
        AtivaBotoes
        If l_opcao = 1 Then
            dta_grau_dependencia.Recordset.AddNew
            AtualTabe
            dta_grau_dependencia.Recordset.Update
        ElseIf l_opcao = 2 Then
            dta_grau_dependencia.Recordset.Edit
            AtualTabe
            dta_grau_dependencia.Recordset.Update
        End If
        l_opcao = 0
        Call BuscaRegistro(l_codigo)
        cmd_novo.SetFocus
    End If
    Exit Sub
FileError:
    ErroArquivo dta_grau_dependencia.Recordset.Name, "Grau de Dependênciao"
    Exit Sub
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If Not Val(txt_codigo) > 0 Then
        MsgBox "Informe o código do grupo.", 64, "Atenção!"
        txt_codigo.SetFocus
    ElseIf Not txt_nome > "" Then
        MsgBox "Informe o nome do grupo.", 64, "Atenção!"
        txt_nome.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_primeiro_Click()
    If dta_grau_dependencia.Recordset.RecordCount > 0 Then
        dta_grau_dependencia.Recordset.MoveFirst
        AtualTela
        cmd_proximo.SetFocus
    End If
End Sub
Private Sub cmd_proximo_Click()
    If dta_grau_dependencia.Recordset.RecordCount > 0 Then
        dta_grau_dependencia.Recordset.MoveNext
        If dta_grau_dependencia.Recordset.EOF Then
            MsgBox "Fim de Arquivo.", 48, "Atenção!"
            dta_grau_dependencia.Recordset.MoveLast
            cmd_anterior.SetFocus
        End If
        AtualTela
    End If
End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub cmd_ultimo_Click()
    If dta_grau_dependencia.Recordset.RecordCount > 0 Then
        dta_grau_dependencia.Recordset.MoveLast
        AtualTela
        cmd_anterior.SetFocus
    End If
End Sub
Private Sub Form_Activate()
    If flag_cadastro_grau_dependencia = 0 Then
        DesativaBotoes
        If BuscaDados Then
            AtivaBotoes
        Else
            cmd_novo.Enabled = True
            cmd_sair.Enabled = True
        End If
        cmd_novo.SetFocus
    Else
        flag_cadastro_grau_dependencia = 0
    End If
End Sub
Private Sub AtivaBotoes()
    cmd_novo.Enabled = True
    cmd_excluir.Enabled = True
    cmd_alterar.Enabled = True
'    cmd_pesquisa.Enabled = True
    cmd_sair.Enabled = True
    tdbgrid_grau_dependencia.Enabled = True
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
    frm_move.Visible = True
End Sub
Private Sub DesativaBotoes()
    cmd_novo.Enabled = False
    cmd_alterar.Enabled = False
    cmd_excluir.Enabled = False
'    cmd_pesquisa.Enabled = False
    cmd_sair.Enabled = False
    tdbgrid_grau_dependencia.Enabled = False
    frm_move.Visible = False
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
End Sub
Function BuscaDados() As Boolean
    BuscaDados = False
    If dta_grau_dependencia.Recordset.RecordCount > 0 Then
        dta_grau_dependencia.Recordset.Seek "<", 9999
        If Not dta_grau_dependencia.Recordset.NoMatch Then
            AtualTela
            BuscaDados = True
        End If
    End If
End Function
Function BuscaRegistro(x_codigo As Integer) As Boolean
    BuscaRegistro = False
    dta_grau_dependencia.Recordset.Seek "=", x_codigo
    If Not dta_grau_dependencia.Recordset.NoMatch Then
        BuscaRegistro = True
        AtualTela
    End If
End Function
Private Sub Form_Deactivate()
    flag_cadastro_grau_dependencia = 1
End Sub
Private Sub Form_Load()
    Screen.MousePointer = 1
    CentraForm Me
    Set dta_grau_dependencia.Recordset = bd_sgp.OpenRecordset("Grau_Dependencia", dbOpenTable)
    dta_grau_dependencia.Recordset.Index = "id_codigo"
    DimensionaGrid
End Sub
Private Sub DimensionaGrid()
    tdbgrid_grau_dependencia.Columns(0).Width = TextWidth(String$(7, "9"))
    tdbgrid_grau_dependencia.Columns(0).HeadAlignment = dbgCenter
    tdbgrid_grau_dependencia.Columns(0).Alignment = dbgRight
    tdbgrid_grau_dependencia.Columns(1).Width = TextWidth(String$(35, "9"))
    tdbgrid_grau_dependencia.Columns(1).HeadAlignment = dbgLeft
    tdbgrid_grau_dependencia.Columns(1).Alignment = dbgLeft
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub tdbgrid_grau_dependencia_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If dta_grau_dependencia.Recordset.RecordCount > 0 And l_opcao <> 1 And tdbgrid_grau_dependencia.Text <> "" Then
        dta_grau_dependencia.Recordset.Seek "=", dta_grau_dependencia.Recordset!Codigo
        If Not dta_grau_dependencia.Recordset.NoMatch Then
            AtualTela
        End If
    End If
End Sub
Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_nome.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_codigo_LostFocus()
    If l_opcao = 1 And txt_codigo <> "" Then
        dta_grau_dependencia.Recordset.Seek "=", Val(txt_codigo)
        If Not dta_grau_dependencia.Recordset.NoMatch Then
            MsgBox "Já existe Grau de Dependência cadastrado com este código." & Chr(10) & Chr(10) & "Mude o código informado.", 64, "Duplicidade de Registro!"
            txt_codigo.SetFocus
        End If
    End If
End Sub
Private Sub txt_nome_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub txt_nome_LostFocus()
    If l_opcao = 1 And txt_nome <> "" Then
        dta_grau_dependencia.Recordset.Index = "id_nome"
        DimensionaGrid
        dta_grau_dependencia.Recordset.Seek ">=", txt_nome, 0
        If Not dta_grau_dependencia.Recordset.NoMatch Then
            If UCase(dta_grau_dependencia.Recordset!Nome) = UCase(txt_nome) Then
                If (MsgBox("Existe outro Grau de Dependência cadastrado com o mesmo nome." & Chr(10) & Chr(10) & "Deseja cadastrar assim mesmo?", 4 + 32 + 256, "Duplicidade de Registro!")) = 7 Then
                    txt_nome.SetFocus
                End If
            End If
        End If
        dta_grau_dependencia.Recordset.Index = "id_codigo"
        DimensionaGrid
    End If
End Sub
