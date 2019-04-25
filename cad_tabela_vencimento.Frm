VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form cadastro_tabela_vencimento 
   Caption         =   "Cadastro da Tabela de Vencimento"
   ClientHeight    =   4935
   ClientLeft      =   630
   ClientTop       =   3315
   ClientWidth     =   6855
   Icon            =   "cad_tabela_vencimento.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "cad_tabela_vencimento.frx":030A
   ScaleHeight     =   4935
   ScaleWidth      =   6855
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   4560
      TabIndex        =   28
      Top             =   3840
      Width           =   2175
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "cad_tabela_vencimento.frx":0750
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "cad_tabela_vencimento.frx":1C4A
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "cad_tabela_vencimento.frx":3144
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "cad_tabela_vencimento.frx":45B6
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   120
      Picture         =   "cad_tabela_vencimento.frx":5B38
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Cria um novo registro."
      Top             =   3960
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   1020
      Picture         =   "cad_tabela_vencimento.frx":71CA
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Altera o registro atual."
      Top             =   3960
      Width           =   795
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   1920
      Picture         =   "cad_tabela_vencimento.frx":86C4
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Exclui o registro atual."
      Top             =   3960
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   2820
      Picture         =   "cad_tabela_vencimento.frx":9D56
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   3960
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Enabled         =   0   'False
      Height          =   1755
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   6615
      Begin VB.TextBox txt_dia_3_f 
         Height          =   285
         Left            =   3240
         MaxLength       =   2
         TabIndex        =   18
         Top             =   1320
         Width           =   435
      End
      Begin VB.TextBox txt_dia_3_i 
         Height          =   285
         Left            =   1020
         MaxLength       =   2
         TabIndex        =   16
         Top             =   1320
         Width           =   435
      End
      Begin VB.TextBox txt_dia_3_v 
         Height          =   285
         Left            =   5820
         MaxLength       =   2
         TabIndex        =   20
         Top             =   1320
         Width           =   435
      End
      Begin VB.TextBox txt_dia_2_v 
         Height          =   285
         Left            =   5820
         MaxLength       =   2
         TabIndex        =   14
         Top             =   960
         Width           =   435
      End
      Begin VB.TextBox txt_dia_2_i 
         Height          =   285
         Left            =   1020
         MaxLength       =   2
         TabIndex        =   10
         Top             =   960
         Width           =   435
      End
      Begin VB.TextBox txt_dia_2_f 
         Height          =   285
         Left            =   3240
         MaxLength       =   2
         TabIndex        =   12
         Top             =   960
         Width           =   435
      End
      Begin VB.TextBox txt_dia_1_v 
         Height          =   285
         Left            =   5820
         MaxLength       =   2
         TabIndex        =   8
         Top             =   600
         Width           =   435
      End
      Begin VB.TextBox txt_dia_1_f 
         Height          =   285
         Left            =   3240
         MaxLength       =   2
         TabIndex        =   6
         Top             =   600
         Width           =   435
      End
      Begin VB.TextBox txt_dia_1_i 
         Height          =   285
         Left            =   1020
         MaxLength       =   2
         TabIndex        =   4
         Top             =   600
         Width           =   435
      End
      Begin VB.TextBox txt_codigo 
         Height          =   285
         Left            =   2160
         MaxLength       =   3
         TabIndex        =   2
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label8 
         Caption         =   "ao dia"
         Height          =   255
         Left            =   2640
         TabIndex        =   17
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Do dia"
         Height          =   255
         Left            =   420
         TabIndex        =   15
         Top             =   1320
         Width           =   555
      End
      Begin VB.Label Label6 
         Caption         =   "vencimento + x dias"
         Height          =   255
         Left            =   4260
         TabIndex        =   19
         Top             =   1320
         Width           =   1515
      End
      Begin VB.Label Label5 
         Caption         =   "vencimento + x dias"
         Height          =   255
         Left            =   4260
         TabIndex        =   13
         Top             =   960
         Width           =   1515
      End
      Begin VB.Label Label2 
         Caption         =   "Do dia"
         Height          =   255
         Left            =   420
         TabIndex        =   9
         Top             =   960
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "ao dia"
         Height          =   255
         Left            =   2640
         TabIndex        =   11
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label15 
         Caption         =   "vencimento + x dias"
         Height          =   255
         Left            =   4260
         TabIndex        =   7
         Top             =   600
         Width           =   1515
      End
      Begin VB.Label Label14 
         Caption         =   "ao dia"
         Height          =   255
         Left            =   2640
         TabIndex        =   5
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Código do Vencimento"
         Height          =   255
         Index           =   2
         Left            =   420
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Do dia"
         Height          =   255
         Left            =   420
         TabIndex        =   3
         Top             =   600
         Width           =   555
      End
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   5940
      Picture         =   "cad_tabela_vencimento.frx":B3E8
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Cancela o registro atual."
      Top             =   3960
      Width           =   795
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   5040
      Picture         =   "cad_tabela_vencimento.frx":C8E2
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Confirma o registro atual."
      Top             =   3960
      Width           =   795
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid 
      Height          =   2055
      Left            =   120
      TabIndex        =   21
      Top             =   1800
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   3625
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   -2147483633
   End
End
Attribute VB_Name = "cadastro_tabela_vencimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lFlagCadastro As Integer
Dim lOpcao As Integer
Dim lCodigo As Integer
Dim lSQl As String
Private rsTabela As New adodb.Recordset
Private Vencimento As New cVencimento
Private Sub AtivaBotoes()
    cmd_novo.Enabled = True
    cmd_alterar.Enabled = True
    cmd_excluir.Enabled = True
    cmd_sair.Enabled = True
    MSFlexGrid.Enabled = True
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
    frm_move.Visible = True
    txt_codigo.Enabled = True
End Sub
Private Sub AtualizaMSFlexGrid()
    Dim i As Integer
    On Error GoTo ErroConsulta
    LimpaMSFlexGrid
    lSQl = "SELECT Codigo, [Dia 1 Inicial], [Dia 1 Final], [Dia 1 Vencimento], [Dia 2 Inicial], [Dia 2 Final], [Dia 2 Vencimento], [Dia 3 Inicial], [Dia 3 Final], [Dia 3 Vencimento]"
    lSQl = lSQl & " FROM Vencimento"
    lSQl = lSQl & " ORDER BY Codigo"
    'Abre RecordSet
    Set rsTabela = New adodb.Recordset
    Set rsTabela = Conectar.RsConexao(lSQl)
    'Verifica movimento
    i = 0
    If rsTabela.RecordCount > 0 Then
        Do Until rsTabela.EOF
            MSFlexGrid.Rows = MSFlexGrid.Rows + 1
            i = i + 1
            MSFlexGrid.Row = i
            MSFlexGrid.Col = 0
            MSFlexGrid.Text = rsTabela("Codigo").Value
            MSFlexGrid.Col = 1
            MSFlexGrid.Text = rsTabela("Dia 1 Inicial").Value
            MSFlexGrid.Col = 2
            MSFlexGrid.Text = rsTabela("Dia 1 Final").Value
            MSFlexGrid.Col = 3
            MSFlexGrid.Text = rsTabela("Dia 1 Vencimento").Value
            MSFlexGrid.Col = 4
            MSFlexGrid.Text = rsTabela("Dia 2 Inicial").Value
            MSFlexGrid.Col = 5
            MSFlexGrid.Text = rsTabela("Dia 2 Final").Value
            MSFlexGrid.Col = 6
            MSFlexGrid.Text = rsTabela("Dia 2 Vencimento").Value
            MSFlexGrid.Col = 7
            MSFlexGrid.Text = rsTabela("Dia 3 Inicial").Value
            MSFlexGrid.Col = 8
            MSFlexGrid.Text = rsTabela("Dia 3 Final").Value
            MSFlexGrid.Col = 9
            MSFlexGrid.Text = rsTabela("Dia 3 Vencimento").Value
            rsTabela.MoveNext
        Loop
    End If
    If rsTabela.State = 1 Then
        rsTabela.Close
    End If
    Exit Sub
    
ErroConsulta:
    If Err = 3075 Then
        MsgBox "Condição inválida.", vbExclamation, "Erro de Consulta"
    Else
        MsgBox Error, vbExclamation, "Erro de Consulta"
    End If
    Exit Sub
End Sub
Private Sub AtualTabe()
    Vencimento.Codigo = Val(txt_codigo)
    Vencimento.Dia1Inicial = Val(txt_dia_1_i)
    Vencimento.Dia1Final = Val(txt_dia_1_f)
    Vencimento.Dia1Vencimento = Val(txt_dia_1_v)
    Vencimento.Dia2Inicial = Val(txt_dia_2_i)
    Vencimento.Dia2Final = Val(txt_dia_2_f)
    Vencimento.Dia2Vencimento = Val(txt_dia_2_v)
    Vencimento.Dia3Inicial = Val(txt_dia_3_i)
    Vencimento.Dia3Final = Val(txt_dia_3_f)
    Vencimento.Dia3Vencimento = Val(txt_dia_3_v)
End Sub
Private Sub AtualTela()
    lCodigo = Vencimento.Codigo
    txt_codigo.Text = Vencimento.Codigo
    txt_dia_1_i.Text = Format(Vencimento.Dia1Inicial, "00")
    txt_dia_1_f.Text = Format(Vencimento.Dia1Final, "00")
    txt_dia_1_v.Text = Format(Vencimento.Dia1Vencimento, "00")
    txt_dia_2_i.Text = Format(Vencimento.Dia2Inicial, "00")
    txt_dia_2_f.Text = Format(Vencimento.Dia2Final, "00")
    txt_dia_2_v.Text = Format(Vencimento.Dia2Vencimento, "00")
    txt_dia_3_i.Text = Format(Vencimento.Dia3Inicial, "00")
    txt_dia_3_f.Text = Format(Vencimento.Dia3Final, "00")
    txt_dia_3_v.Text = Format(Vencimento.Dia3Vencimento, "00")
    frm_dados.Enabled = False
End Sub
Private Sub DesativaBotoes()
    cmd_novo.Enabled = False
    cmd_alterar.Enabled = False
    cmd_excluir.Enabled = False
    cmd_sair.Enabled = False
    MSFlexGrid.Enabled = False
    frm_move.Visible = False
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
End Sub
Private Sub Finaliza()
    Set Vencimento = Nothing
    frm_cadastro.Show
End Sub
Private Sub Inclui()
    lOpcao = 1
    DesativaBotoes
    cmd_novo.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    txt_codigo.Text = 0
    If Vencimento.LocalizarUltimo Then
        txt_codigo.Text = Vencimento.Codigo
    End If
    txt_codigo.Text = Format(Val(txt_codigo.Text) + 1, "00")
End Sub

Private Sub cmd_alterar_Click()
    lOpcao = 2
    DesativaBotoes
    cmd_alterar.Visible = True
    cmd_alterar.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    frm_dados.Enabled = True
    txt_codigo.Enabled = False
    txt_dia_1_i.SetFocus
End Sub
Private Sub cmd_anterior_Click()
    If Vencimento.LocalizarAnterior Then
        AtualTela
    Else
        MsgBox "Início de Arquivo.", vbInformation, "Atenção!"
        cmd_proximo.SetFocus
    End If
End Sub
Private Sub cmd_cancelar_Click()
    LimpaTela
    lOpcao = 0
    If Vencimento.LocalizarCodigo(lCodigo) Then
        AtualTela
        AtivaBotoes
        cmd_alterar.SetFocus
    Else
        DesativaBotoes
        cmd_novo.Enabled = True
        cmd_sair.Enabled = True
        cmd_novo.SetFocus
    End If
End Sub
Private Sub LimpaMSFlexGrid()
    Dim i As Integer
    MSFlexGrid.WordWrap = True
    MSFlexGrid.Cols = 10
    MSFlexGrid.Rows = 2
    MSFlexGrid.Row = 1
    For i = 0 To (MSFlexGrid.Cols - 1)
        MSFlexGrid.Col = i
        MSFlexGrid.Text = ""
    Next
    MSFlexGrid.RowHeight(0) = 500
    MSFlexGrid.Row = 0
    i = 0
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Código"
    MSFlexGrid.ColWidth(i) = 700
    MSFlexGrid.ColAlignment(i) = flexAlignGeneral
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Dia 1 Inicial"
    MSFlexGrid.ColWidth(i) = 600
    MSFlexGrid.ColAlignment(i) = flexAlignGeneral
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Dia 1 Final"
    MSFlexGrid.ColWidth(i) = 600
    MSFlexGrid.ColAlignment(i) = flexAlignGeneral
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Dia 1 Vencimento"
    MSFlexGrid.ColWidth(i) = 1000
    MSFlexGrid.ColAlignment(i) = flexAlignGeneral
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Dia 2 Inicial"
    MSFlexGrid.ColWidth(i) = 600
    MSFlexGrid.ColAlignment(i) = flexAlignGeneral
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Dia 2 Final"
    MSFlexGrid.ColWidth(i) = 600
    MSFlexGrid.ColAlignment(i) = flexAlignGeneral
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Dia 2 Vencimento"
    MSFlexGrid.ColWidth(i) = 1000
    MSFlexGrid.ColAlignment(i) = flexAlignGeneral
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Dia 3 Inicial"
    MSFlexGrid.ColWidth(i) = 600
    MSFlexGrid.ColAlignment(i) = flexAlignGeneral
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Dia 3 Final"
    MSFlexGrid.ColWidth(i) = 600
    MSFlexGrid.ColAlignment(i) = flexAlignGeneral
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Dia 3 Vencimento"
    MSFlexGrid.ColWidth(i) = 1000
    MSFlexGrid.ColAlignment(i) = flexAlignGeneral
End Sub
Private Sub LimpaTela()
    txt_codigo.Text = ""
    txt_dia_1_i.Text = ""
    txt_dia_1_f.Text = ""
    txt_dia_1_v.Text = ""
    txt_dia_2_i.Text = ""
    txt_dia_2_f.Text = ""
    txt_dia_2_v.Text = ""
    txt_dia_3_i.Text = ""
    txt_dia_3_f.Text = ""
    txt_dia_3_v.Text = ""
End Sub
Private Sub cmd_excluir_Click()
    If Val(txt_codigo.Text) > 0 Then
        If (MsgBox("Deseja excluir este registro?", 4 + 32 + 256, "Exclusão de Registro!")) = 6 Then
            If Vencimento.Excluir(txt_codigo.Text) Then
                LimpaTela
                If Vencimento.LocalizarUltimo Then
                    AtualTela
                    AtualizaMSFlexGrid
                Else
                    DesativaBotoes
                    cmd_novo.Enabled = True
                    cmd_sair.Enabled = True
                    cmd_novo.SetFocus
                End If
            Else
                MsgBox "Não foi possivel excluir este registro!", vbInformation, "Erro de Verificação!"
            End If
        End If
    End If
End Sub
Private Sub cmd_novo_Click()
    LimpaTela
    Inclui
    frm_dados.Enabled = True
    txt_dia_1_i.SetFocus
End Sub
Private Sub cmd_ok_Click()
    On Error GoTo FileError
    If ValidaCampos Then
        AtivaBotoes
        If lOpcao = 1 Then
            AtualTabe
            If Vencimento.Incluir Then
                lCodigo = Val(txt_codigo.Text)
            Else
                MsgBox "Não foi possível incluir este registro!", vbInformation, "Erro de Verificação!"
            End If
        ElseIf lOpcao = 2 Then
            AtualTabe
            If Not Vencimento.Alterar(txt_codigo.Text) Then
                MsgBox "Não foi possível alterar este registro!", vbInformation, "Erro de Verificação!"
            End If
        End If
        If Vencimento.LocalizarCodigo(lCodigo) Then
            AtualTela
            AtualizaMSFlexGrid
        End If
        lOpcao = 0
        cmd_novo.SetFocus
    End If
    Exit Sub
FileError:
    MsgBox Error
    'ErroArquivo dta_vencimento.Recordset.Name, "Vencimentoo"
    Exit Sub
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If Not Val(txt_codigo) > 0 Then
        MsgBox "Informe o código da tabela.", vbInformation, "Atenção!"
        txt_codigo.SetFocus
    ElseIf Not Val(txt_dia_1_i) > 0 Then
        MsgBox "Informe o dia inicial (1).", vbInformation, "Atenção!"
        txt_dia_1_i.SetFocus
    ElseIf Not Val(txt_dia_1_f) > 0 Then
        MsgBox "Informe o dia final (1).", vbInformation, "Atenção!"
        txt_dia_1_f.SetFocus
    ElseIf Not Val(txt_dia_1_v) > 0 Then
        MsgBox "Informe o dia do vencimento (1).", vbInformation, "Atenção!"
        txt_dia_1_v.SetFocus
    ElseIf Not Val(txt_dia_2_i) > 0 Then
        MsgBox "Informe o dia inicial (2).", vbInformation, "Atenção!"
        txt_dia_2_i.SetFocus
    ElseIf Not Val(txt_dia_2_f) > 0 Then
        MsgBox "Informe o dia final (2).", vbInformation, "Atenção!"
        txt_dia_2_f.SetFocus
    ElseIf Not Val(txt_dia_2_v) > 0 Then
        MsgBox "Informe o dia do vencimento (2).", vbInformation, "Atenção!"
        txt_dia_2_v.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_primeiro_Click()
    If Vencimento.LocalizarPrimeiro Then
        AtualTela
        cmd_proximo.SetFocus
    Else
        MsgBox "Não há registro.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    If Vencimento.LocalizarProximo Then
        AtualTela
    Else
        MsgBox "Fim de Arquivo.", vbInformation, "Atenção!"
        cmd_anterior.SetFocus
    End If
End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub cmd_ultimo_Click()
    If Vencimento.LocalizarUltimo Then
        AtualTela
        cmd_anterior.SetFocus
    Else
        MsgBox "Não há registro.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub Form_Activate()
    If lFlagCadastro = 0 Then
        DesativaBotoes
        If Vencimento.LocalizarUltimo Then
            AtualTela
            AtivaBotoes
        Else
            cmd_novo.Enabled = True
            cmd_sair.Enabled = True
        End If
        cmd_novo.SetFocus
    Else
        lFlagCadastro = 0
    End If
End Sub
Private Sub Form_Deactivate()
    lFlagCadastro = 1
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 And lOpcao = 0 Then
        KeyCode = 0
        cmd_novo_Click
    ElseIf KeyCode = vbKeyF3 And lOpcao = 0 Then
        KeyCode = 0
        cmd_alterar_Click
    ElseIf KeyCode = vbKeyF4 And Shift = 0 Then
        KeyCode = 0
        cmd_excluir_Click
    ElseIf KeyCode = vbKeyF7 And lOpcao = 0 Then
        KeyCode = 0
        cmd_primeiro_Click
    ElseIf KeyCode = vbKeyF8 And lOpcao = 0 Then
        KeyCode = 0
        cmd_anterior_Click
    ElseIf KeyCode = vbKeyF9 And lOpcao = 0 Then
        KeyCode = 0
        cmd_proximo_Click
    ElseIf KeyCode = vbKeyF10 And lOpcao = 0 Then
        KeyCode = 0
        cmd_ultimo_Click
    ElseIf KeyCode = vbKeyF11 And lOpcao > 0 Then
        KeyCode = 0
        cmd_ok_Click
    ElseIf KeyCode = vbKeyF12 And lOpcao > 0 Then
        KeyCode = 0
        cmd_cancelar_Click
    End If
End Sub
Private Sub Form_Load()
    Screen.MousePointer = 1
    CentraForm Me
    AtualizaMSFlexGrid
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub MSFlexGrid_RowColChange()
    If lOpcao = 0 And MSFlexGrid.TextMatrix(MSFlexGrid.Row, 0) <> "" Then
        lCodigo = Val(MSFlexGrid.TextMatrix(MSFlexGrid.Row, 0))
        If Vencimento.LocalizarCodigo(lCodigo) Then
            AtualTela
        End If
    End If
End Sub
Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_dia_1_i.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_codigo_LostFocus()
    If lOpcao = 1 And txt_codigo.Text <> "" Then
        If Vencimento.LocalizarCodigo(Val(txt_codigo.Text)) Then
            MsgBox "Já existe vencimento cadastrado com este código." & Chr(10) & Chr(10) & "Mude o código informado.", vbInformation, "Duplicidade de Registro!"
            txt_codigo.SetFocus
        End If
    End If
End Sub
Private Sub txt_dia_1_f_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_dia_1_v.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_dia_1_i_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_dia_1_f.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_dia_1_v_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_dia_2_i.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_dia_2_f_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_dia_2_v.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_dia_2_i_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_dia_2_f.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_dia_2_v_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_dia_3_i.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_dia_3_f_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_dia_3_v.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_dia_3_i_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_dia_3_f.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_dia_3_v_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
