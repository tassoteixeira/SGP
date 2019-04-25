VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form cadastro_tabela_folha 
   Caption         =   "Cadastro da Tabela da Folha"
   ClientHeight    =   6375
   ClientLeft      =   120
   ClientTop       =   1875
   ClientWidth     =   7035
   Icon            =   "cad_tabela_folha.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "cad_tabela_folha.frx":030A
   ScaleHeight     =   6375
   ScaleWidth      =   7035
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   120
      Picture         =   "cad_tabela_folha.frx":0750
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Cria um novo registro."
      Top             =   5400
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   1020
      Picture         =   "cad_tabela_folha.frx":1DE2
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Altera o registro atual."
      Top             =   5400
      Width           =   795
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   1920
      Picture         =   "cad_tabela_folha.frx":32DC
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "Exclui o registro atual."
      Top             =   5400
      Width           =   795
   End
   Begin VB.CommandButton cmd_pesquisa 
      Caption         =   "&Pesquisa"
      Height          =   855
      Left            =   2820
      Picture         =   "cad_tabela_folha.frx":496E
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "Pesquisa um registro específico."
      Top             =   5400
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3720
      Picture         =   "cad_tabela_folha.frx":5DE0
      Style           =   1  'Graphical
      TabIndex        =   46
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   5400
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Enabled         =   0   'False
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6795
      Begin VB.TextBox txt_salario_familia_2 
         Height          =   285
         Left            =   5580
         MaxLength       =   10
         TabIndex        =   12
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txt_acima_de 
         Height          =   285
         Left            =   2820
         MaxLength       =   10
         TabIndex        =   10
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txt_observacao_3 
         Height          =   285
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   39
         Top             =   4860
         Width           =   5355
      End
      Begin VB.TextBox txt_observacao_2 
         Height          =   285
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   37
         Top             =   4500
         Width           =   5355
      End
      Begin VB.TextBox txt_observacao_1 
         Height          =   285
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   35
         Top             =   4140
         Width           =   5355
      End
      Begin VB.TextBox txt_valor_4_b 
         Height          =   285
         Left            =   3780
         MaxLength       =   10
         TabIndex        =   32
         Top             =   3660
         Width           =   1095
      End
      Begin VB.TextBox txt_valor_4_a 
         Height          =   285
         Left            =   1620
         MaxLength       =   10
         TabIndex        =   31
         Top             =   3660
         Width           =   1095
      End
      Begin VB.TextBox txt_porcentagem_4 
         Height          =   285
         Left            =   5760
         MaxLength       =   6
         TabIndex        =   33
         Top             =   3660
         Width           =   555
      End
      Begin VB.TextBox txt_valor_3_b 
         Height          =   285
         Left            =   3780
         MaxLength       =   10
         TabIndex        =   28
         Top             =   3240
         Width           =   1095
      End
      Begin VB.TextBox txt_valor_3_a 
         Height          =   285
         Left            =   1620
         MaxLength       =   10
         TabIndex        =   27
         Top             =   3240
         Width           =   1095
      End
      Begin VB.TextBox txt_porcentagem_3 
         Height          =   285
         Left            =   5760
         MaxLength       =   6
         TabIndex        =   29
         Top             =   3240
         Width           =   555
      End
      Begin VB.TextBox txt_valor_2_b 
         Height          =   285
         Left            =   3780
         MaxLength       =   10
         TabIndex        =   24
         Top             =   2820
         Width           =   1095
      End
      Begin VB.TextBox txt_valor_2_a 
         Height          =   285
         Left            =   1620
         MaxLength       =   10
         TabIndex        =   23
         Top             =   2820
         Width           =   1095
      End
      Begin VB.TextBox txt_porcentagem_2 
         Height          =   285
         Left            =   5760
         MaxLength       =   6
         TabIndex        =   25
         Top             =   2820
         Width           =   555
      End
      Begin VB.TextBox txt_valor_1_b 
         Height          =   285
         Left            =   3780
         MaxLength       =   10
         TabIndex        =   20
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox txt_valor_1_a 
         Height          =   285
         Left            =   1620
         MaxLength       =   10
         TabIndex        =   19
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox txt_porcentagem_1 
         Height          =   285
         Left            =   5760
         MaxLength       =   6
         TabIndex        =   21
         Top             =   2400
         Width           =   555
      End
      Begin VB.TextBox txt_cesta_basica 
         Height          =   285
         Left            =   5580
         MaxLength       =   10
         TabIndex        =   6
         Top             =   420
         Width           =   1095
      End
      Begin VB.TextBox txt_salario_familia_1 
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   8
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txt_salario_minimo 
         Height          =   285
         Left            =   2820
         MaxLength       =   10
         TabIndex        =   4
         Top             =   420
         Width           =   1095
      End
      Begin MSMask.MaskEdBox msk_data 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   420
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   7
         Format          =   "mm/yyyy"
         Mask            =   "##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         Caption         =   "Salário &família II"
         Height          =   255
         Index           =   14
         Left            =   5580
         TabIndex        =   11
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label Label3 
         Caption         =   "Acima de"
         Height          =   255
         Index           =   13
         Left            =   2820
         TabIndex        =   9
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Line Line12 
         X1              =   120
         X2              =   6660
         Y1              =   1740
         Y2              =   1740
      End
      Begin VB.Label Label3 
         Caption         =   "Observação 3"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   38
         Top             =   4860
         Width           =   1155
      End
      Begin VB.Label Label3 
         Caption         =   "Observação 2"
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   36
         Top             =   4500
         Width           =   1155
      End
      Begin VB.Label Label3 
         Caption         =   "Observação 1"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   34
         Top             =   4140
         Width           =   1155
      End
      Begin VB.Line Line5 
         X1              =   6660
         X2              =   6660
         Y1              =   1740
         Y2              =   4020
      End
      Begin VB.Label Label3 
         Caption         =   "Faixa &4"
         Height          =   255
         Index           =   9
         Left            =   300
         TabIndex        =   30
         Top             =   3720
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Faixa &3"
         Height          =   255
         Index           =   8
         Left            =   300
         TabIndex        =   26
         Top             =   3300
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Faixa &2"
         Height          =   255
         Index           =   7
         Left            =   300
         TabIndex        =   22
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Faixas"
         Height          =   195
         Index           =   4
         Left            =   300
         TabIndex        =   14
         Top             =   2100
         Width           =   615
      End
      Begin VB.Line Line11 
         X1              =   6660
         X2              =   120
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Line Line10 
         X1              =   6660
         X2              =   120
         Y1              =   3180
         Y2              =   3180
      End
      Begin VB.Line Line9 
         X1              =   6660
         X2              =   120
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line Line4 
         X1              =   120
         X2              =   6660
         Y1              =   2340
         Y2              =   2340
      End
      Begin VB.Line Line8 
         X1              =   5400
         X2              =   5400
         Y1              =   2040
         Y2              =   4020
      End
      Begin VB.Line Line7 
         X1              =   3240
         X2              =   3240
         Y1              =   2040
         Y2              =   4020
      End
      Begin VB.Line Line6 
         X1              =   1080
         X2              =   1080
         Y1              =   2040
         Y2              =   4020
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   120
         Y1              =   1740
         Y2              =   4020
      End
      Begin VB.Line Line3 
         X1              =   120
         X2              =   6660
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   6660
         Y1              =   4020
         Y2              =   4020
      End
      Begin VB.Label Label3 
         Caption         =   "Faixa &1"
         Height          =   255
         Index           =   3
         Left            =   300
         TabIndex        =   18
         Top             =   2460
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Porcentagem"
         Height          =   255
         Index           =   17
         Left            =   5520
         TabIndex        =   17
         Top             =   2100
         Width           =   1035
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Valor final"
         Height          =   255
         Index           =   16
         Left            =   3840
         TabIndex        =   16
         Top             =   2100
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Valor inicial"
         Height          =   255
         Index           =   0
         Left            =   1680
         TabIndex        =   15
         Top             =   2100
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Cesta &básica"
         Height          =   255
         Index           =   6
         Left            =   5580
         TabIndex        =   5
         Top             =   180
         Width           =   1035
      End
      Begin VB.Label Label3 
         Caption         =   "Salário &família I"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label Label3 
         Caption         =   "Mês/Ano da &tabela"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   1
         Top             =   180
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Salário &mínimo"
         Height          =   255
         Index           =   1
         Left            =   2820
         TabIndex        =   3
         Top             =   180
         Width           =   1155
      End
      Begin VB.Label Label2 
         Caption         =   "I.N.S.S."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   2880
         TabIndex        =   13
         Top             =   1740
         Width           =   975
      End
   End
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   4740
      TabIndex        =   47
      Top             =   5280
      Width           =   2175
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "cad_tabela_folha.frx":7472
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "cad_tabela_folha.frx":89F4
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "cad_tabela_folha.frx":9E66
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "cad_tabela_folha.frx":B360
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   5220
      Picture         =   "cad_tabela_folha.frx":C85A
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "Confirma o registro atual."
      Top             =   5400
      Width           =   795
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   6120
      Picture         =   "cad_tabela_folha.frx":DE64
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "Cancela o registro atual."
      Top             =   5400
      Width           =   795
   End
End
Attribute VB_Name = "cadastro_tabela_folha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lFlagCadastro As Integer
Dim lOpcao As Integer
Dim lData As String
Private TabelaFolha As New CTabelaFolha
Private Sub AtivaBotoes()
    cmd_novo.Enabled = True
    cmd_alterar.Enabled = True
    cmd_pesquisa.Enabled = True
    cmd_excluir.Enabled = True
    cmd_sair.Enabled = True
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
    frm_move.Visible = True
    msk_data.Enabled = True
End Sub
Private Sub AtualTabe()
    TabelaFolha.MesAno = Mid(msk_data, 4, 4) & Mid(msk_data, 1, 2)
    TabelaFolha.SalarioMinimo = fValidaValor2(txt_salario_minimo)
    TabelaFolha.CestaBasica = fValidaValor2(txt_cesta_basica)
    TabelaFolha.SalarioFamilia1 = fValidaValor2(txt_salario_familia_1)
    TabelaFolha.SalarioFamiliaAcimade = fValidaValor2(txt_acima_de)
    TabelaFolha.SalarioFamilia2 = fValidaValor2(txt_salario_familia_2)
    TabelaFolha.ValorInicial1 = fValidaValor2(txt_valor_1_a)
    TabelaFolha.ValorFinal1 = fValidaValor2(txt_valor_1_b)
    TabelaFolha.Percentual1 = fValidaValor2(txt_porcentagem_1)
    TabelaFolha.ValorInicial2 = fValidaValor2(txt_valor_2_a)
    TabelaFolha.ValorFinal2 = fValidaValor2(txt_valor_2_b)
    TabelaFolha.Percentual2 = fValidaValor2(txt_porcentagem_2)
    TabelaFolha.ValorInicial3 = fValidaValor2(txt_valor_3_a)
    TabelaFolha.ValorFinal3 = fValidaValor2(txt_valor_3_b)
    TabelaFolha.Percentual3 = fValidaValor2(txt_porcentagem_3)
    TabelaFolha.ValorInicial4 = fValidaValor2(txt_valor_4_a)
    TabelaFolha.ValorFinal4 = fValidaValor2(txt_valor_4_b)
    TabelaFolha.Percentual4 = fValidaValor2(txt_porcentagem_4)
    TabelaFolha.Observacao1 = txt_observacao_1
    TabelaFolha.Observacao2 = txt_observacao_2
    TabelaFolha.Observacao3 = txt_observacao_3
End Sub
Private Sub AtualTela()
    lData = TabelaFolha.MesAno
    msk_data = Mid(TabelaFolha.MesAno, 5, 2) & "/" & Mid(TabelaFolha.MesAno, 1, 4)
    txt_salario_minimo = Format(TabelaFolha.SalarioMinimo, "###,##0.00")
    txt_cesta_basica = Format(TabelaFolha.CestaBasica, "###,##0.00")
    txt_salario_familia_1 = Format(TabelaFolha.SalarioFamilia1, "###,##0.00")
    txt_acima_de = Format(TabelaFolha.SalarioFamiliaAcimade, "###,##0.00")
    txt_salario_familia_2 = Format(TabelaFolha.SalarioFamilia2, "###,##0.00")
    txt_valor_1_a = Format(TabelaFolha.ValorInicial1, "###,##0.00")
    txt_valor_1_b = Format(TabelaFolha.ValorFinal1, "###,##0.00")
    txt_porcentagem_1 = Format(TabelaFolha.Percentual1, "##0.00")
    txt_valor_2_a = Format(TabelaFolha.ValorInicial2, "###,##0.00")
    txt_valor_2_b = Format(TabelaFolha.ValorFinal2, "###,##0.00")
    txt_porcentagem_2 = Format(TabelaFolha.Percentual2, "##0.00")
    txt_valor_3_a = Format(TabelaFolha.ValorInicial3, "###,##0.00")
    txt_valor_3_b = Format(TabelaFolha.ValorFinal3, "###,##0.00")
    txt_porcentagem_3 = Format(TabelaFolha.Percentual3, "##0.00")
    txt_valor_4_a = Format(TabelaFolha.ValorInicial4, "###,##0.00")
    txt_valor_4_b = Format(TabelaFolha.ValorFinal4, "###,##0.00")
    txt_porcentagem_4 = Format(TabelaFolha.Percentual4, "##0.00")
    txt_observacao_1 = TabelaFolha.Observacao1
    txt_observacao_2 = TabelaFolha.Observacao2
    txt_observacao_3 = TabelaFolha.Observacao3
    frm_dados.Enabled = False
End Sub
Private Sub DesativaBotoes()
    cmd_novo.Enabled = False
    cmd_alterar.Enabled = False
    cmd_pesquisa.Enabled = False
    cmd_excluir.Enabled = False
    cmd_sair.Enabled = False
    frm_move.Visible = False
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
End Sub
Private Sub Finaliza()
    Set TabelaFolha = Nothing
    frm_cadastro.Show
End Sub
Private Sub Inclui()
    Dim x_data As Date
    lOpcao = 1
    DesativaBotoes
    cmd_novo.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    msk_data.Text = Format(g_data_def, "mm") & "/" & Format(g_data_def, "yyyy")
    If TabelaFolha.LocalizarUltimo Then
        x_data = CDate("01/" & Mid(TabelaFolha.MesAno, 5, 2) & "/" & Mid(TabelaFolha.MesAno, 1, 4)) + 35
        msk_data = Format(x_data, "mm") & "/" & Format(x_data, "yyyy")
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
    msk_data.Enabled = False
    txt_salario_minimo.SetFocus
End Sub
Private Sub cmd_anterior_Click()
    If TabelaFolha.LocalizarAnterior Then
        AtualTela
    Else
        MsgBox "Início de Arquivo.", vbInformation, "Atenção!"
        cmd_proximo.SetFocus
    End If
End Sub
Private Sub cmd_cancelar_Click()
    LimpaTela
    lOpcao = 0
    If TabelaFolha.LocalizarCodigo(lData) Then
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
Private Sub LimpaTela()
    msk_data = "__/____"
    txt_salario_minimo = ""
    txt_cesta_basica = ""
    txt_salario_familia_1 = ""
    txt_acima_de = ""
    txt_salario_familia_2 = ""
    txt_valor_1_a = ""
    txt_valor_1_b = ""
    txt_porcentagem_1 = ""
    txt_valor_2_a = ""
    txt_valor_2_b = ""
    txt_porcentagem_2 = ""
    txt_valor_3_a = ""
    txt_valor_3_b = ""
    txt_porcentagem_3 = ""
    txt_valor_4_a = ""
    txt_valor_4_b = ""
    txt_porcentagem_4 = ""
    txt_observacao_1 = ""
    txt_observacao_2 = ""
    txt_observacao_3 = ""
End Sub
Private Sub cmd_excluir_Click()
    Dim xData As String
    xData = Mid(msk_data, 4, 4) & Mid(msk_data, 1, 2)
    If xData <> "" Then
        If (MsgBox("Deseja Realmente Excluir Este Registro?", 4 + 32 + 256, "Exclusão de Registro!")) = 6 Then
            If TabelaFolha.Excluir(xData) Then
                LimpaTela
                If TabelaFolha.LocalizarUltimo Then
                    AtualTela
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
    If (MsgBox("Deseja aproveitar os dados deste registro?", 4 + 32 + 0, "Aproveitamento de Dados!")) = 7 Then
        LimpaTela
    End If
    Inclui
    frm_dados.Enabled = True
    txt_salario_minimo.SetFocus
End Sub
Private Sub cmd_ok_Click()
    On Error GoTo FileError
    If ValidaCampos Then
        AtivaBotoes
        If lOpcao = 1 Then
            AtualTabe
            If TabelaFolha.Incluir Then
                lData = Mid(msk_data, 4, 4) & Mid(msk_data, 1, 2)
            Else
                MsgBox "Não foi possível incluir este registro!", vbInformation, "Erro de Verificação!"
            End If
        ElseIf lOpcao = 2 Then
            AtualTabe
            If Not TabelaFolha.Alterar(lData) Then
                MsgBox "Não foi possível alterar este registro!", vbInformation, "Erro de Verificação!"
            End If
        End If
        lOpcao = 0
        Call TabelaFolha.LocalizarCodigo(lData)
        AtualTela
        cmd_novo.SetFocus
    End If
    Exit Sub
FileError:
    MsgBox Error
    'ErroArquivo tbl_tabela_folha.Name, "Tabela da Folhaa"
    Exit Sub
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If Val(Mid(msk_data, 1, 2)) < 1 Or Val(Mid(msk_data, 1, 2)) > 14 Then
        MsgBox "Informe o mês entre 01 a 13.", vbInformation, "Atenção!"
        msk_data.SetFocus
    ElseIf Val(Mid(msk_data, 4, 4)) < 1998 Or Val(Mid(msk_data, 4, 4)) > 2500 Then
        MsgBox "Informe o ano entre 1998 a 2500.", vbInformation, "Atenção!"
        msk_data.SetFocus
    ElseIf Not Val(txt_salario_minimo) > 0 Then
        MsgBox "Informe o valor do salário mínimo.", vbInformation, "Atenção!"
        txt_salario_minimo.SetFocus
    ElseIf Not Val(txt_cesta_basica) > 0 Then
        MsgBox "Informe o valor da cesta básica.", vbInformation, "Atenção!"
        txt_cesta_basica.SetFocus
    ElseIf Not Val(txt_salario_familia_1) > 0 Then
        MsgBox "Informe o valor do salário família I.", vbInformation, "Atenção!"
        txt_salario_familia_1.SetFocus
    ElseIf Not Val(txt_acima_de) > 0 Then
        MsgBox "Informe o valor para cálculo do salário família.", vbInformation, "Atenção!"
        txt_acima_de.SetFocus
    ElseIf Not Val(txt_salario_familia_2) > 0 Then
        MsgBox "Informe o valor do salário família II.", vbInformation, "Atenção!"
        txt_salario_familia_2.SetFocus
    ElseIf Val(txt_valor_1_a) <> 0 Then
        MsgBox "O valor inicial da faixa 1 deve ser 0.", vbInformation, "Atenção!"
        txt_valor_1_a.SetFocus
    ElseIf Not Val(txt_valor_1_b) > 0 Then
        MsgBox "Informe o valor final da faixa 1.", vbInformation, "Atenção!"
        txt_valor_1_b.SetFocus
    ElseIf Not Val(txt_porcentagem_1) > 0 Then
        MsgBox "Informe a porcentagem da faixa 1.", vbInformation, "Atenção!"
        txt_porcentagem_1.SetFocus
    ElseIf Not fValidaValor2(txt_valor_2_a) > fValidaValor2(txt_valor_1_b) Then
        MsgBox "O valor inicial da faixa 2 deve ser maior que " & txt_valor_1_b & ".", vbInformation, "Atenção!"
        txt_valor_2_a.SetFocus
    ElseIf Not fValidaValor2(txt_valor_2_b) > fValidaValor2(txt_valor_2_a) Then
        MsgBox "O valor final da faixa 2 deve ser maior que " & txt_valor_2_a & ".", vbInformation, "Atenção!"
        txt_valor_2_b.SetFocus
    ElseIf Not fValidaValor2(txt_porcentagem_2) > fValidaValor2(txt_porcentagem_1) Then
        MsgBox "A porcentagem da faixa 2 deve ser maior que " & txt_porcentagem_1 & ".", vbInformation, "Atenção!"
        txt_porcentagem_2.SetFocus
    ElseIf Not fValidaValor2(txt_valor_3_a) > fValidaValor2(txt_valor_2_b) Then
        MsgBox "O valor inicial da faixa 3 deve ser maior que " & txt_valor_2_b & ".", vbInformation, "Atenção!"
        txt_valor_3_a.SetFocus
    ElseIf Not fValidaValor2(txt_valor_3_b) > fValidaValor2(txt_valor_3_a) Then
        MsgBox "O valor final da faixa 3 deve ser maior que " & txt_valor_3_a & ".", vbInformation, "Atenção!"
        txt_valor_3_b.SetFocus
    ElseIf Not fValidaValor2(txt_porcentagem_3) > fValidaValor2(txt_porcentagem_2) Then
        MsgBox "A porcentagem da faixa 3 deve ser maior que " & txt_porcentagem_2 & ".", vbInformation, "Atenção!"
        txt_porcentagem_3.SetFocus
    ElseIf Not fValidaValor2(txt_valor_4_a) > fValidaValor2(txt_valor_3_b) Then
        MsgBox "O valor inicial da faixa 4 deve ser maior que " & txt_valor_3_b & ".", vbInformation, "Atenção!"
        txt_valor_4_a.SetFocus
    ElseIf Not fValidaValor2(txt_valor_4_b) > fValidaValor2(txt_valor_4_a) Then
        MsgBox "O valor final da faixa 4 deve ser maior que " & txt_valor_4_a & ".", vbInformation, "Atenção!"
        txt_valor_4_b.SetFocus
    ElseIf Not fValidaValor2(txt_porcentagem_4) > fValidaValor2(txt_porcentagem_3) Then
        MsgBox "A porcentagem da faixa 4 deve ser maior que " & txt_porcentagem_3 & " .", vbInformation, "Atenção!"
        txt_porcentagem_4.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_pesquisa_Click()
    consulta_tabela_folha.Show 1
    If Len(g_string) > 0 Then
        lData = RetiraGString(1)
        If TabelaFolha.LocalizarCodigo(lData) Then
            AtualTela
        End If
    End If
End Sub
Private Sub cmd_primeiro_Click()
    If TabelaFolha.LocalizarPrimeiro Then
        AtualTela
        cmd_proximo.SetFocus
    Else
        MsgBox "Não há registro.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    If TabelaFolha.LocalizarProximo Then
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
    If TabelaFolha.LocalizarUltimo Then
        AtualTela
        cmd_anterior.SetFocus
    Else
        MsgBox "Não há registro.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub Form_Activate()
    If lFlagCadastro = 0 Then
        DesativaBotoes
        If TabelaFolha.LocalizarUltimo Then
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
    ElseIf KeyCode = vbKeyF5 And lOpcao = 0 Then
        KeyCode = 0
        cmd_pesquisa_Click
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
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0
        txt_salario_minimo.SetFocus
    End If
End Sub
Private Sub msk_data_LostFocus()
    Dim xData As String
    xData = Mid(msk_data.Text, 4, 4) & Mid(msk_data.Text, 1, 2)
    If lOpcao = 1 And msk_data.Text <> "" Then
        If TabelaFolha.LocalizarCodigo(xData) Then
            MsgBox "Já existe tabela da folha cadastrada nesta data." & Chr(10) & Chr(10) & "Mude a data informada.", vbInformation, "Duplicidade de Registro!"
            msk_data.SetFocus
        End If
    End If
End Sub
Private Sub txt_acima_de_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_salario_familia_2.SetFocus
    End If
End Sub
Private Sub txt_acima_de_LostFocus()
    txt_acima_de = Format(txt_acima_de, "###,##0.00")
End Sub
Private Sub txt_cesta_basica_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_salario_familia_1.SetFocus
    End If
End Sub
Private Sub txt_cesta_basica_LostFocus()
    txt_cesta_basica = Format(txt_cesta_basica, "###,##0.00")
End Sub
Private Sub txt_observacao_1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_observacao_2.SetFocus
    End If
End Sub
Private Sub txt_observacao_2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_observacao_3.SetFocus
    End If
End Sub
Private Sub txt_observacao_3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub txt_porcentagem_1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_valor_2_a.SetFocus
    End If
End Sub
Private Sub txt_porcentagem_1_LostFocus()
    txt_porcentagem_1 = Format(txt_porcentagem_1, "##0.00")
End Sub
Private Sub txt_porcentagem_2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_valor_3_a.SetFocus
    End If
End Sub
Private Sub txt_porcentagem_2_LostFocus()
    txt_porcentagem_2 = Format(txt_porcentagem_2, "##0.00")
End Sub
Private Sub txt_porcentagem_3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_valor_4_a.SetFocus
    End If
End Sub
Private Sub txt_porcentagem_3_LostFocus()
    txt_porcentagem_3 = Format(txt_porcentagem_3, "##0.00")
End Sub
Private Sub txt_porcentagem_4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_observacao_1.SetFocus
    End If
End Sub
Private Sub txt_porcentagem_4_LostFocus()
    txt_porcentagem_4 = Format(txt_porcentagem_4, "##0.00")
End Sub
Private Sub txt_salario_familia_1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_acima_de.SetFocus
    End If
End Sub
Private Sub txt_salario_familia_1_LostFocus()
    txt_salario_familia_1 = Format(txt_salario_familia_1, "###,##0.00")
End Sub
Private Sub txt_salario_familia_2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_valor_1_a.SetFocus
    End If
End Sub
Private Sub txt_salario_familia_2_LostFocus()
    txt_salario_familia_2 = Format(txt_salario_familia_2, "###,##0.00")
End Sub
Private Sub txt_salario_minimo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_cesta_basica.SetFocus
    End If
End Sub
Private Sub txt_salario_minimo_LostFocus()
    txt_salario_minimo = Format(txt_salario_minimo, "###,##0.00")
End Sub
Private Sub txt_valor_1_a_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_valor_1_b.SetFocus
    End If
End Sub
Private Sub txt_valor_1_a_LostFocus()
    txt_valor_1_a = Format(txt_valor_1_a, "###,##0.00")
End Sub
Private Sub txt_valor_1_b_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_porcentagem_1.SetFocus
    End If
End Sub
Private Sub txt_valor_1_b_LostFocus()
    txt_valor_1_b = Format(txt_valor_1_b, "###,##0.00")
End Sub
Private Sub txt_valor_2_a_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_valor_2_b.SetFocus
    End If
End Sub
Private Sub txt_valor_2_a_LostFocus()
    txt_valor_2_a = Format(txt_valor_2_a, "###,##0.00")
End Sub
Private Sub txt_valor_2_b_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_porcentagem_2.SetFocus
    End If
End Sub
Private Sub txt_valor_2_b_LostFocus()
    txt_valor_2_b = Format(txt_valor_2_b, "###,##0.00")
End Sub
Private Sub txt_valor_3_a_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_valor_3_b.SetFocus
    End If
End Sub
Private Sub txt_valor_3_a_LostFocus()
    txt_valor_3_a = Format(txt_valor_3_a, "###,##0.00")
End Sub
Private Sub txt_valor_3_b_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_porcentagem_3.SetFocus
    End If
End Sub
Private Sub txt_valor_3_b_LostFocus()
    txt_valor_3_b = Format(txt_valor_3_b, "###,##0.00")
End Sub
Private Sub txt_valor_4_a_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_valor_4_b.SetFocus
    End If
End Sub
Private Sub txt_valor_4_a_LostFocus()
    txt_valor_4_a = Format(txt_valor_4_a, "###,##0.00")
End Sub
Private Sub txt_valor_4_b_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_porcentagem_4.SetFocus
    End If
End Sub
Private Sub txt_valor_4_b_LostFocus()
    txt_valor_4_b = Format(txt_valor_4_b, "###,##0.00")
End Sub
