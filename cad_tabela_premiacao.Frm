VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form cadastro_tabela_premiacao 
   Caption         =   "Cadastro da Tabela de Premiação"
   ClientHeight    =   5295
   ClientLeft      =   630
   ClientTop       =   3315
   ClientWidth     =   6870
   Icon            =   "cad_tabela_premiacao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "cad_tabela_premiacao.frx":030A
   ScaleHeight     =   5295
   ScaleWidth      =   6870
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   120
      Picture         =   "cad_tabela_premiacao.frx":0750
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Cria um novo registro."
      Top             =   4320
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   1020
      Picture         =   "cad_tabela_premiacao.frx":1DE2
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Altera o registro atual."
      Top             =   4320
      Width           =   795
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   1920
      Picture         =   "cad_tabela_premiacao.frx":32DC
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Exclui o registro atual."
      Top             =   4320
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   2820
      Picture         =   "cad_tabela_premiacao.frx":496E
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   4320
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Enabled         =   0   'False
      Height          =   2115
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   6615
      Begin VB.TextBox txt_premiacao_3 
         Height          =   285
         Left            =   5820
         MaxLength       =   3
         TabIndex        =   16
         Top             =   1680
         Width           =   555
      End
      Begin VB.TextBox txt_percentual_3 
         Height          =   285
         Left            =   2340
         MaxLength       =   3
         TabIndex        =   14
         Top             =   1680
         Width           =   555
      End
      Begin VB.TextBox txt_premiacao_2 
         Height          =   285
         Left            =   5820
         MaxLength       =   3
         TabIndex        =   12
         Top             =   1320
         Width           =   555
      End
      Begin VB.TextBox txt_percentual_2 
         Height          =   285
         Left            =   2340
         MaxLength       =   3
         TabIndex        =   10
         Top             =   1320
         Width           =   555
      End
      Begin VB.TextBox txt_premiacao_1 
         Height          =   285
         Left            =   5820
         MaxLength       =   3
         TabIndex        =   8
         Top             =   960
         Width           =   555
      End
      Begin VB.TextBox txt_percentual_1 
         Height          =   285
         Left            =   2340
         MaxLength       =   3
         TabIndex        =   6
         Top             =   960
         Width           =   555
      End
      Begin VB.TextBox msk_valor 
         Height          =   285
         Left            =   2340
         MaxLength       =   10
         TabIndex        =   4
         Top             =   600
         Width           =   1095
      End
      Begin MSMask.MaskEdBox msk_mes_ano 
         Height          =   315
         Left            =   2340
         TabIndex        =   2
         Top             =   180
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   7
         Format          =   "mm/yyyy"
         Mask            =   "##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         Caption         =   "Premiação de "
         Height          =   255
         Index           =   7
         Left            =   4680
         TabIndex        =   15
         Top             =   1680
         Width           =   1035
      End
      Begin VB.Label Label3 
         Caption         =   "Percentual sobre valor base"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   13
         Top             =   1680
         Width           =   1995
      End
      Begin VB.Label Label3 
         Caption         =   "Premiação de "
         Height          =   255
         Index           =   5
         Left            =   4680
         TabIndex        =   11
         Top             =   1320
         Width           =   1035
      End
      Begin VB.Label Label3 
         Caption         =   "Percentual sobre valor base"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Width           =   1995
      End
      Begin VB.Label Label3 
         Caption         =   "Premiação de "
         Height          =   255
         Index           =   3
         Left            =   4680
         TabIndex        =   7
         Top             =   960
         Width           =   1035
      End
      Begin VB.Label Label3 
         Caption         =   "Percentual sobre valor base"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   1995
      End
      Begin VB.Label Label3 
         Caption         =   "Valor base p/ premiação"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   1995
      End
      Begin VB.Label Label3 
         Caption         =   "Mês/Ano da tabela"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1995
      End
   End
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   4560
      TabIndex        =   24
      Top             =   4200
      Width           =   2175
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "cad_tabela_premiacao.frx":6000
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "cad_tabela_premiacao.frx":74FA
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "cad_tabela_premiacao.frx":89F4
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "cad_tabela_premiacao.frx":9E66
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   5940
      Picture         =   "cad_tabela_premiacao.frx":B3E8
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Cancela o registro atual."
      Top             =   4320
      Width           =   795
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   5040
      Picture         =   "cad_tabela_premiacao.frx":C8E2
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Confirma o registro atual."
      Top             =   4320
      Width           =   795
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid 
      Height          =   2055
      Left            =   120
      TabIndex        =   17
      Top             =   2160
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   3625
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   -2147483633
   End
End
Attribute VB_Name = "cadastro_tabela_premiacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lFlagCadastro As Integer
Dim lOpcao As Integer
Dim lData As Date
Dim lSQl As String
Private rsTabela As New adodb.Recordset
Private Premiacao As New cPremiacao
Private Sub AtivaBotoes()
    cmd_novo.Enabled = True
    cmd_alterar.Enabled = True
    cmd_excluir.Enabled = True
    cmd_sair.Enabled = True
    MSFlexGrid.Enabled = True
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
    frm_move.Visible = True
    msk_mes_ano.Enabled = True
End Sub
Private Sub AtualizaMSFlexGrid()
    Dim i As Integer
    On Error GoTo ErroConsulta
    LimpaMSFlexGrid
    lSQl = "SELECT [Mes Ano], [Valor Base], [Percentual Base 1], [Percentual Comissao 1], [Percentual Base 2], [Percentual Comissao 2], [Percentual Base 3], [Percentual Comissao 3]"
    lSQl = lSQl & " FROM Tabela_Premiacao"
    lSQl = lSQl & " WHERE Empresa = " & g_empresa
    lSQl = lSQl & " ORDER BY [Mes Ano]"
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
            MSFlexGrid.Text = Mid(rsTabela("Mes Ano").Value, 4, 7)
            MSFlexGrid.Col = 1
            MSFlexGrid.Text = Format(rsTabela("Valor Base").Value, "####,###,##0.00")
            MSFlexGrid.Col = 2
            MSFlexGrid.Text = rsTabela("Percentual Base 1").Value
            MSFlexGrid.Col = 3
            MSFlexGrid.Text = rsTabela("Percentual Comissao 1").Value
            MSFlexGrid.Col = 4
            MSFlexGrid.Text = rsTabela("Percentual Base 2").Value
            MSFlexGrid.Col = 5
            MSFlexGrid.Text = rsTabela("Percentual Comissao 2").Value
            MSFlexGrid.Col = 6
            MSFlexGrid.Text = rsTabela("Percentual Base 3").Value
            MSFlexGrid.Col = 7
            MSFlexGrid.Text = rsTabela("Percentual Comissao 3").Value
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
    Premiacao.Empresa = g_empresa
    Premiacao.MesAno = CDate("01/" & msk_mes_ano)
    Premiacao.ValorBase = fValidaValor2(msk_valor)
    Premiacao.PercentualBase1 = Val(txt_percentual_1)
    Premiacao.PercentualComissao1 = Val(txt_premiacao_1)
    Premiacao.PercentualBase2 = Val(txt_percentual_2)
    Premiacao.PercentualComissao2 = Val(txt_premiacao_2)
    Premiacao.PercentualBase3 = Val(txt_percentual_3)
    Premiacao.PercentualComissao3 = Val(txt_premiacao_3)
    lData = Premiacao.MesAno
End Sub
Private Sub AtualTela()
    lData = Premiacao.MesAno
    msk_mes_ano = Format(Premiacao.MesAno, "mm") & "/" & Format(Premiacao.MesAno, "yyyy")
    msk_valor = Format(Premiacao.ValorBase, "###,##0.00")
    txt_percentual_1.Text = Format(Premiacao.PercentualBase1, "##0")
    txt_premiacao_1.Text = Format(Premiacao.PercentualComissao1, "##0")
    txt_percentual_2.Text = Format(Premiacao.PercentualBase2, "##0")
    txt_premiacao_2.Text = Format(Premiacao.PercentualComissao2, "##0")
    txt_percentual_3.Text = Format(Premiacao.PercentualBase3, "##0")
    txt_premiacao_3.Text = Format(Premiacao.PercentualComissao3, "##0")
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
    Set Premiacao = Nothing
    frm_cadastro.Show
End Sub
Private Sub Inclui()
    lOpcao = 1
    DesativaBotoes
    cmd_novo.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    msk_mes_ano.Text = Format(g_data_def, "mm") & "/" & Format(g_data_def, "yyyy")
    If Premiacao.LocalizarUltimo(g_empresa) Then
       msk_mes_ano.Text = Format(CDate(Premiacao.MesAno) + 35, "mm") & "/" & Format(CDate(Premiacao.MesAno) + 35, "yyyy")
    End If
End Sub
Private Sub cmd_alterar_Click()
    lOpcao = 2
    DesativaBotoes
    cmd_alterar.Visible = True
    cmd_alterar.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    frm_dados.Enabled = True
    msk_mes_ano.Enabled = False
    msk_valor.SetFocus
End Sub
Private Sub cmd_anterior_Click()
    If Premiacao.LocalizarAnterior Then
        AtualTela
    Else
        MsgBox "Início de Arquivo.", vbInformation, "Atenção!"
        cmd_proximo.SetFocus
    End If
End Sub
Private Sub cmd_cancelar_Click()
    LimpaTela
    lOpcao = 0
    If Premiacao.LocalizarCodigo(g_empresa, lData) Then
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
    MSFlexGrid.Cols = 8
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
    MSFlexGrid.Text = "Mês Ano"
    MSFlexGrid.ColWidth(i) = 700
    MSFlexGrid.ColAlignment(i) = flexAlignGeneral
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Valor Base"
    MSFlexGrid.ColWidth(i) = 800
    MSFlexGrid.ColAlignment(i) = flexAlignGeneral
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Percentual Base 1"
    MSFlexGrid.ColWidth(i) = 900
    MSFlexGrid.ColAlignment(i) = flexAlignGeneral
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Percentual Comissao 1"
    MSFlexGrid.ColWidth(i) = 900
    MSFlexGrid.ColAlignment(i) = flexAlignGeneral
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Percentual Base 2"
    MSFlexGrid.ColWidth(i) = 900
    MSFlexGrid.ColAlignment(i) = flexAlignGeneral
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Percentual Comissao 2"
    MSFlexGrid.ColWidth(i) = 900
    MSFlexGrid.ColAlignment(i) = flexAlignGeneral
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Percentual Base 3"
    MSFlexGrid.ColWidth(i) = 900
    MSFlexGrid.ColAlignment(i) = flexAlignGeneral
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Percentual Comissao 3"
    MSFlexGrid.ColWidth(i) = 900
    MSFlexGrid.ColAlignment(i) = flexAlignGeneral
End Sub
Private Sub LimpaTela()
    msk_mes_ano.Text = "__/____"
    msk_valor.Text = ""
    txt_percentual_1.Text = ""
    txt_premiacao_1.Text = ""
    txt_percentual_2.Text = ""
    txt_premiacao_2.Text = ""
    txt_percentual_3.Text = ""
    txt_premiacao_3.Text = ""
End Sub
Private Sub cmd_excluir_Click()
    If Val(msk_mes_ano.Text) > 0 Then
        If (MsgBox("Deseja excluir este registro?", 4 + 32 + 256, "Exclusão de Registro!")) = 6 Then
            If Premiacao.Excluir(g_empresa, msk_mes_ano.Text) Then
                LimpaTela
                If Premiacao.LocalizarUltimo(g_empresa) Then
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
    msk_mes_ano.SetFocus
End Sub
Private Sub cmd_ok_Click()
    On Error GoTo FileError
    If ValidaCampos Then
        AtivaBotoes
        If lOpcao = 1 Then
            AtualTabe
            If Premiacao.Incluir Then
                lData = "01/" & msk_mes_ano.Text
            Else
                MsgBox "Não foi possível incluir este registro!", vbInformation, "Erro de Verificação!"
            End If
        ElseIf lOpcao = 2 Then
            AtualTabe
            If Not Premiacao.Alterar(g_empresa, CDate("01/" & msk_mes_ano.Text)) Then
                MsgBox "Não foi possível alterar este registro!", vbInformation, "Erro de Verificação!"
            End If
        End If
        If Premiacao.LocalizarCodigo(g_empresa, lData) Then
            AtualTela
            AtualizaMSFlexGrid
        End If
        lOpcao = 0
        cmd_novo.SetFocus
    End If
    Exit Sub
FileError:
    MsgBox Error
    'ErroArquivo dta_tabela_premiacao.Recordset.Name, "Vencimentoo"
    Exit Sub
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    Dim x_data As String
    x_data = "01/" & Mid(msk_mes_ano, 1, 2) & "/" & Mid(msk_mes_ano, 4, 4)
    If Not IsDate(x_data) Then
        MsgBox "Informe o mês/ano da tabela.", vbInformation, "Atenção!"
        msk_mes_ano.SetFocus
    ElseIf Not fValidaValor2(msk_valor) > 0 Then
        MsgBox "Informe o valor base p/ premiação.", vbInformation, "Atenção!"
        msk_valor.SetFocus
    ElseIf Not Val(txt_percentual_1) > 0 Then
        MsgBox "Informe o percentual sobre valor base.", vbInformation, "Atenção!"
        txt_percentual_1.SetFocus
    ElseIf Not Val(txt_premiacao_1) > 0 Then
        MsgBox "Informe a premiação.", vbInformation, "Atenção!"
        txt_premiacao_1.SetFocus
    ElseIf Not Val(txt_percentual_2) > 0 Then
        MsgBox "Informe o percentual sobre valor base.", vbInformation, "Atenção!"
        txt_percentual_2.SetFocus
    ElseIf Not Val(txt_premiacao_2) > 0 Then
        MsgBox "Informe a premiação.", vbInformation, "Atenção!"
        txt_premiacao_2.SetFocus
    ElseIf Not Val(txt_percentual_3) > 0 Then
        MsgBox "Informe o percentual sobre valor base.", vbInformation, "Atenção!"
        txt_percentual_3.SetFocus
    ElseIf Not Val(txt_premiacao_3) > 0 Then
        MsgBox "Informe a premiação.", vbInformation, "Atenção!"
        txt_premiacao_3.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_primeiro_Click()
    If Premiacao.LocalizarPrimeiro Then
        AtualTela
        cmd_proximo.SetFocus
    Else
        MsgBox "Não há registro.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    If Premiacao.LocalizarProximo Then
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
    If Premiacao.LocalizarUltimo(g_empresa) Then
        AtualTela
        cmd_anterior.SetFocus
    Else
        MsgBox "Não há registro.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub Form_Activate()
    If lFlagCadastro = 0 Then
        DesativaBotoes
        If Premiacao.LocalizarUltimo(g_empresa) Then
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
    If lOpcao = 0 And IsDate("01/" & MSFlexGrid.TextMatrix(MSFlexGrid.Row, 0)) Then
        lData = "01/" & MSFlexGrid.TextMatrix(MSFlexGrid.Row, 0)
        If Premiacao.LocalizarCodigo(g_empresa, lData) Then
            AtualTela
        End If
    End If
End Sub
Private Sub msk_mes_ano_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_valor.SetFocus
    End If
End Sub
Private Sub msk_mes_ano_LostFocus()
    Dim x_data As String
    If lOpcao > 0 Then
        x_data = "01/" & msk_mes_ano
        If lOpcao = 1 And IsDate(x_data) Then
            If Premiacao.LocalizarCodigo(g_empresa, CDate(x_data)) Then
                MsgBox "Já existe tabela de premiação cadastrada nesta data." & Chr(10) & Chr(10) & "Mude a data informada.", vbInformation, "Duplicidade de Registro!"
                msk_mes_ano.SetFocus
            End If
        End If
    End If
End Sub
Private Sub msk_valor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_percentual_1.SetFocus
    End If
End Sub
Private Sub msk_valor_LostFocus()
    If Val(msk_valor) > 0 Then
        msk_valor = Format(msk_valor, "###,##0.00")
    End If
End Sub
Private Sub txt_percentual_1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_premiacao_1.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_percentual_1_LostFocus()
    If Val(txt_percentual_1) > 0 Then
        txt_percentual_1 = Format(txt_percentual_1, "##0")
    End If
End Sub
Private Sub txt_percentual_2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_premiacao_2.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_percentual_2_LostFocus()
    If Val(txt_percentual_2) > 0 Then
        txt_percentual_2 = Format(txt_percentual_2, "##0")
    End If
End Sub
Private Sub txt_percentual_3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_premiacao_3.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_percentual_3_LostFocus()
    If Val(txt_percentual_3) > 0 Then
        txt_percentual_3 = Format(txt_percentual_3, "##0")
    End If
End Sub
Private Sub txt_premiacao_1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_percentual_2.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_premiacao_1_LostFocus()
    If Val(txt_premiacao_1) > 0 Then
        txt_premiacao_1 = Format(txt_premiacao_1, "##0")
    End If
End Sub
Private Sub txt_premiacao_2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_percentual_3.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_premiacao_2_LostFocus()
    If Val(txt_premiacao_2) > 0 Then
        txt_premiacao_2 = Format(txt_premiacao_2, "##0")
    End If
End Sub
Private Sub txt_premiacao_3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_premiacao_3_LostFocus()
    If Val(txt_premiacao_3) > 0 Then
        txt_premiacao_3 = Format(txt_premiacao_3, "##0")
    End If
End Sub
