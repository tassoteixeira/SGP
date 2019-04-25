VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form movimento_folha 
   Caption         =   "Movimento da Folha de Pagamento"
   ClientHeight    =   6075
   ClientLeft      =   165
   ClientTop       =   285
   ClientWidth     =   7875
   Icon            =   "movimento_folha.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "movimento_folha.frx":030A
   ScaleHeight     =   6075
   ScaleWidth      =   7875
   Begin VB.CommandButton cmd_sair 
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3720
      Picture         =   "movimento_folha.frx":0750
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   5100
      Width           =   795
   End
   Begin VB.CommandButton cmd_pesquisa 
      Caption         =   "&Pesquisa"
      Height          =   855
      Left            =   2820
      Picture         =   "movimento_folha.frx":1DE2
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Pesquisa um registro específico."
      Top             =   5100
      Width           =   795
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   1920
      Picture         =   "movimento_folha.frx":3254
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Exclui o registro atual."
      Top             =   5100
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   1020
      Picture         =   "movimento_folha.frx":48E6
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Altera o registro atual."
      Top             =   5100
      Width           =   795
   End
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   120
      Picture         =   "movimento_folha.frx":5DE0
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Cria um novo registro."
      Top             =   5100
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   4620
      Picture         =   "movimento_folha.frx":7472
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Imprime contra-cheque do funcionário."
      Top             =   5100
      Width           =   795
   End
   Begin VB.Frame frmDados 
      Enabled         =   0   'False
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   7635
      Begin VB.CommandButton cmd_processamento 
         Caption         =   "&Processamento"
         Height          =   735
         Left            =   6000
         Picture         =   "movimento_folha.frx":874C
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Processa folha de pagamento automaticamente."
         Top             =   1260
         Width           =   1515
      End
      Begin VB.Data dta_provento_desconto 
         Caption         =   "dta_provento_desconto"
         Connect         =   "Access"
         DatabaseName    =   "Sgp_data.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   4860
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Tabela_Provento_Desconto"
         Top             =   900
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox txt_provento_desconto 
         Height          =   300
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   7
         Top             =   900
         Width           =   555
      End
      Begin VB.Data dta_funcionario 
         Caption         =   "dta_funcionario"
         Connect         =   "Access"
         DatabaseName    =   "Sgp_data.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   4860
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Funcionario"
         Top             =   540
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox txt_funcionario 
         Height          =   300
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   4
         Top             =   540
         Width           =   555
      End
      Begin VB.TextBox txt_quantidade 
         Height          =   285
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   10
         Top             =   1260
         Width           =   1095
      End
      Begin VB.TextBox txt_valor 
         Height          =   285
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   12
         Top             =   1620
         Width           =   1095
      End
      Begin MSDBCtls.DBCombo dbcbo_funcionario 
         Bindings        =   "movimento_folha.frx":9A26
         Height          =   315
         Left            =   2640
         TabIndex        =   5
         Top             =   540
         Width           =   4875
         _ExtentX        =   8599
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSMask.MaskEdBox msk_data 
         Height          =   300
         Left            =   2040
         TabIndex        =   2
         Top             =   180
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   7
         Format          =   "mm/yyyy"
         Mask            =   "##/####"
         PromptChar      =   "_"
      End
      Begin MSDBCtls.DBCombo dbcbo_provento_desconto 
         Bindings        =   "movimento_folha.frx":9A44
         Height          =   315
         Left            =   2640
         TabIndex        =   8
         Top             =   900
         Width           =   4875
         _ExtentX        =   8599
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         Caption         =   "Provento/&Desconto"
         Height          =   315
         Index           =   7
         Left            =   120
         TabIndex        =   6
         Top             =   900
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "&Funcionário"
         Height          =   315
         Index           =   6
         Left            =   120
         TabIndex        =   3
         Top             =   540
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "&Mês/Ano"
         Height          =   300
         Index           =   3
         Left            =   120
         TabIndex        =   1
         Top             =   180
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "&Quantidade"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   1260
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "&Valor"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   1620
         Width           =   1815
      End
   End
   Begin TabDlg.SSTab tab_dados 
      Height          =   2835
      Left            =   120
      TabIndex        =   13
      Top             =   2160
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   5001
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "P&roventos/Descontos"
      TabPicture(0)   =   "movimento_folha.frx":9A68
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl_liquido"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "grid_provento_desconto"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin MSGrid.Grid grid_provento_desconto 
         Height          =   2115
         Left            =   60
         TabIndex        =   14
         Top             =   360
         Width           =   7515
         _Version        =   65536
         _ExtentX        =   13256
         _ExtentY        =   3731
         _StockProps     =   77
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Cols            =   4
         FixedCols       =   0
         ScrollBars      =   2
      End
      Begin VB.Label lbl_liquido 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6180
         TabIndex        =   16
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Total Líquido a Receber"
         Height          =   255
         Index           =   0
         Left            =   4260
         TabIndex        =   15
         Top             =   2520
         Width           =   1815
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   5580
      TabIndex        =   25
      Top             =   4980
      Width           =   2175
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "movimento_folha.frx":9A84
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "movimento_folha.frx":B006
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "movimento_folha.frx":C478
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "movimento_folha.frx":D972
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   6060
      Picture         =   "movimento_folha.frx":EE6C
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Confirma o registro atual."
      Top             =   5100
      Width           =   795
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   6960
      Picture         =   "movimento_folha.frx":10476
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Cancela o registro atual."
      Top             =   5100
      Width           =   795
   End
End
Attribute VB_Name = "movimento_folha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flag_movimento_folha As Integer
Dim lOpcao As String
Dim l_empresa As Integer
Dim l_ano_mes As String * 6
Dim l_funcionario As Integer
Dim l_funcionario2 As Integer
Dim l_codigo As Integer
Dim l_gravados As Long
Dim l_local As Integer
Dim tbl_dependente As Table
Dim tbl_empresa As Table
Dim tbl_funcionario As Table
Dim tbl_movimento_falta_caixa As Table
Dim tbl_movimento_falta_funcionario As Table
Dim tbl_movimento_folha As Table
Dim tbl_tabela_folha As Table
Dim tbl_tabela_provento_desconto As Table
Function AceitaProventoDesconto(x_codigo As Integer, x_funcionario As Integer) As Boolean
    AceitaProventoDesconto = True
    'Adicional de Periculosidade
    If x_codigo = 5 And tbl_funcionario![Adicional de Periculosidade] = False Then
        AceitaProventoDesconto = False
    'Adicional Noturno
    ElseIf x_codigo = 10 And tbl_funcionario![Adicional Noturno] = False Then
        AceitaProventoDesconto = False
    'Salário Família
    ElseIf x_codigo = 15 Then
        If CalculaDependente(x_funcionario, "S") = 0 Then
            AceitaProventoDesconto = False
        End If
    'Adiantamento 1a Parcela 13o Salario
    ElseIf x_codigo = 505 Then
        If Val(Mid(l_ano_mes, 5, 2)) <> 14 Then
            AceitaProventoDesconto = False
        End If
    'Faltas
    ElseIf x_codigo = 510 Then
        If QuantidadeFaltas(x_funcionario, "N") = 0 Then
            AceitaProventoDesconto = False
        End If
    'D.S.R. Remunerado
    ElseIf x_codigo = 515 Then
        If QuantidadeFaltas(x_funcionario, "S") = 0 Then
            AceitaProventoDesconto = False
        End If
    'Vale Transporte
    ElseIf x_codigo = 530 And tbl_funcionario![Vale Transporte] = False Then
        AceitaProventoDesconto = False
    'Seguro de Vida
    ElseIf x_codigo = 550 And tbl_funcionario![Seguro de Vida] = False Then
        AceitaProventoDesconto = False
    'Cesta Básica
    ElseIf x_codigo = 560 And tbl_funcionario![Cesta Basica] = False Then
        AceitaProventoDesconto = False
    'Contribuição Assistencial
    ElseIf x_codigo = 610 And tbl_tabela_provento_desconto!Automatico = False Then
        AceitaProventoDesconto = False
    End If
End Function
Private Sub AdcionaDadosGridProventoDesconto(x_codigo As Integer, x_nome As String, x_provento As Currency, x_desconto As Currency)
    Dim x_i As Integer
    grid_provento_desconto.Row = grid_provento_desconto.Rows - 1
    grid_provento_desconto.Col = 0
    grid_provento_desconto.Text = Format(x_codigo, "#000")
    grid_provento_desconto.Col = 1
    grid_provento_desconto.Text = x_nome
    grid_provento_desconto.Col = 2
    If x_provento > 0 Then
        grid_provento_desconto.Text = Format(x_provento, "###,##0.00") & "  "
    End If
    grid_provento_desconto.Col = 3
    If x_desconto > 0 Then
        grid_provento_desconto.Text = Format(x_desconto, "###,##0.00") & "  "
    End If
    grid_provento_desconto.Rows = grid_provento_desconto.Rows + 1
End Sub
Private Sub AtivaBotoes()
    cmd_novo.Enabled = True
    cmd_alterar.Enabled = True
    cmd_excluir.Enabled = True
    cmd_pesquisa.Enabled = True
    cmd_imprimir.Enabled = True
    cmd_sair.Enabled = True
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
    frm_move.Visible = True
End Sub
Private Sub Inclui()
    lOpcao = 1
    DesativaBotoes
    cmd_novo.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    If Val(l_ano_mes) > 0 Then
        msk_data = Mid(l_ano_mes, 5, 2) & "/" & Mid(l_ano_mes, 1, 4)
    Else
        msk_data = Format(g_data_def, "mm") & "/" & Format(g_data_def, "yyyy")
    End If
End Sub
Private Sub AtualTabe()
    l_ano_mes = Mid(msk_data, 4, 4) & Mid(msk_data, 1, 2)
    l_funcionario = Val(txt_funcionario)
    l_codigo = Val(txt_provento_desconto)
    With tbl_movimento_folha
        !Empresa = g_empresa
        ![Mes Ano] = Mid(msk_data, 4, 4) & Mid(msk_data, 1, 2)
        ![Codigo do Funcionario] = Val(txt_funcionario)
        ![Codigo do Movimento] = Val(txt_provento_desconto)
        !Quantidade = fValidaValor2(txt_quantidade)
        !valor = fValidaValor2(txt_valor)
        !Status = "          "
    End With
End Sub
Private Sub AtualTela()
    With tbl_movimento_folha
        l_ano_mes = Mid(![Mes Ano], 1, 4) & Mid(![Mes Ano], 5, 2)
        l_funcionario = ![Codigo do Funcionario]
        tbl_funcionario.Seek "=", g_empresa, ![Codigo do Funcionario]
        tbl_tabela_folha.Seek "=", l_ano_mes
        l_codigo = ![Codigo do Movimento]
        msk_data = Mid(![Mes Ano], 5, 2) & "/" & Mid(![Mes Ano], 1, 4)
        txt_funcionario = ![Codigo do Funcionario]
        dbcbo_funcionario.BoundText = ![Codigo do Funcionario]
        txt_provento_desconto = ![Codigo do Movimento]
        dbcbo_provento_desconto.BoundText = ![Codigo do Movimento]
        txt_quantidade = Format(!Quantidade, "###,##0.00")
        txt_valor = Format(!valor, "###,##0.00")
        MontaGridProventoDesconto
    End With
    frmDados.Enabled = False
    'tab_dados.Enabled = False
End Sub
Function BuscaRegistro(x_data As String, x_funcionario As Integer, x_codigo As Integer) As Boolean
    BuscaRegistro = False
    With tbl_movimento_folha
        If .RecordCount > 0 Then
            .Seek "=", g_empresa, x_data, x_funcionario, x_codigo
            If Not .NoMatch Then
                AtualTela
                BuscaRegistro = True
                Exit Function
            End If
        End If
    End With
End Function
Function BuscaRegistro2(x_data As String, x_funcionario As Integer, x_codigo As Integer) As Boolean
    BuscaRegistro2 = False
    With tbl_movimento_folha
        If .RecordCount > 0 Then
            .Seek "=", g_empresa, x_data, x_funcionario, x_codigo
            If Not .NoMatch Then
                BuscaRegistro2 = True
                Exit Function
            End If
        End If
    End With
End Function
Function BuscaDados() As Boolean
    BuscaDados = False
    With tbl_movimento_folha
        If .RecordCount > 0 Then
            If lOpcao = 3 Then
                If Not .EOF Then
                    .MoveNext
                    If Not .EOF Then
                        If !Empresa = g_empresa Then
                            AtualTela
                            BuscaDados = True
                            Exit Function
                        End If
                    End If
                End If
            End If
            .Seek "<", g_empresa, CDate("31/12/2500"), 9999, 9999
            If Not .NoMatch Then
                If !Empresa = g_empresa Then
                    AtualTela
                    BuscaDados = True
                    Exit Function
                End If
            End If
        End If
        l_gravados = 0
        LimpaTela
    End With
End Function
Function Calcula13Salario() As Integer
    Dim x_data As Date
    Dim x_data2 As Date
    Calcula13Salario = 0
    If Not IsDate(tbl_funcionario![Data de Admissao]) Then
        MsgBox "Funcionario.: " & tbl_funcionario!Nome
        Exit Function
    End If
    x_data = tbl_funcionario![Data de Admissao]
    If Year(x_data) < Val(Mid(l_ano_mes, 1, 4)) Then
        Calcula13Salario = 12
        Exit Function
    Else
        If Day(x_data) <= 14 Then
            Calcula13Salario = 13 - Month(x_data)
            Exit Function
        End If
    End If
    x_data2 = x_data
    Do Until Month(x_data2) <> Month(x_data)
        x_data2 = x_data2 + 1
    Loop
    x_data2 = x_data2 - 1
    If (x_data2 - x_data) + 1 >= 15 Then
        Calcula13Salario = 13 - Month(x_data)
        Exit Function
    Else
        x_data = x_data2 + 1
        Calcula13Salario = 13 - Month(x_data)
        Exit Function
    End If
End Function
Function CalculaBaseCalculo(x_codigo As Integer) As Currency
    CalculaBaseCalculo = 0
    If Len(Trim(tbl_tabela_provento_desconto![Base para Calculo])) > 0 Then
        CalculaBaseCalculo = CalculaBaseCalculoComposta
    End If
    'Salário Base
    If x_codigo = 1 Then
        CalculaBaseCalculo = tbl_funcionario![Salario Base]
    'Salário Família
    ElseIf x_codigo = 15 Then
        If CalculaBaseCalculo > tbl_tabela_folha![Salario Familia Acima De] Then
            CalculaBaseCalculo = tbl_tabela_folha![Salario Familia 2]
        Else
            CalculaBaseCalculo = tbl_tabela_folha![Salario Familia 1]
        End If
    'Adiantamento 1a Parcela 13o Salario
    ElseIf x_codigo = 505 Then
'        If BuscaRegistro2(Mid(l_ano_mes, 1, 4) & "13", l_funcionario2, 1) Then
'            CalculaBaseCalculo = tbl_movimento_folha!Valor
'        End If
'        If BuscaRegistro2(Mid(l_ano_mes, 1, 4) & "13", l_funcionario2, 5) Then
'            CalculaBaseCalculo = CalculaBaseCalculo + tbl_movimento_folha!Valor
'        End If
        If BuscaRegistro2(Mid(l_ano_mes, 1, 4) & "14", l_funcionario2, 1) Then
            CalculaBaseCalculo = tbl_movimento_folha!valor / 2
        End If
        If BuscaRegistro2(Mid(l_ano_mes, 1, 4) & "14", l_funcionario2, 5) Then
            CalculaBaseCalculo = CalculaBaseCalculo + (tbl_movimento_folha!valor / 2)
        End If
    'Cesta Básica
    ElseIf x_codigo = 560 Then
        CalculaBaseCalculo = tbl_tabela_folha![Cesta Basica]
    End If
    CalculaBaseCalculo = Format(CalculaBaseCalculo, "########0.00")
    If lOpcao = 2 Then
        Call BuscaRegistro2(l_ano_mes, l_funcionario, l_codigo)
    End If
End Function
Function CalculaBaseCalculoComposta() As Currency
    Dim x_codigo As Integer
    Dim i As Integer
    Dim i2 As Integer
    CalculaBaseCalculoComposta = 0
    i2 = Len(tbl_tabela_provento_desconto![Base para Calculo]) / 5
    i = 1
    Do Until i > i2
        x_codigo = RetiraString(tbl_tabela_provento_desconto![Base para Calculo], i)
        If BuscaRegistro2(Mid(msk_data, 4, 4) & Mid(msk_data, 1, 2), l_funcionario, x_codigo) Then
            CalculaBaseCalculoComposta = CalculaBaseCalculoComposta + tbl_movimento_folha!valor
        End If
        i = i + 1
    Loop
    'Se o movimento for INSS
    If tbl_tabela_provento_desconto!Codigo = 520 Then
        'Verifica se tem "Falta" e deduz na base de cálculo
        If BuscaRegistro2(Mid(msk_data, 4, 4) & Mid(msk_data, 1, 2), l_funcionario, 510) Then
            CalculaBaseCalculoComposta = CalculaBaseCalculoComposta - tbl_movimento_folha!valor
        End If
        'Verifica se tem "DSR" e deduz na base de cálculo
        If BuscaRegistro2(Mid(msk_data, 4, 4) & Mid(msk_data, 1, 2), l_funcionario, 515) Then
            CalculaBaseCalculoComposta = CalculaBaseCalculoComposta - tbl_movimento_folha!valor
        End If
    End If
    CalculaBaseCalculoComposta = Format(CalculaBaseCalculoComposta, "########0.00")
End Function
Function CalculaDependente(x_funcionario As Integer, x_salario_familia As String) As Integer
    Dim x_data As Date
    Dim x_data_nascimento As Date
    Dim x_mes As Integer
    CalculaDependente = 0
    If Mid(l_ano_mes, 5, 2) > 12 Then
        Exit Function
    End If
    x_data = "28/" & Mid(l_ano_mes, 5, 2) & "/" & Mid(l_ano_mes, 1, 4)
    x_mes = Month(x_data)
    Do Until x_mes <> Month(x_data)
        x_data = x_data + 1
    Loop
    x_data = x_data - 1
    With tbl_dependente
        If .RecordCount > 0 Then
            .Seek ">=", g_empresa, x_funcionario, 0
            If Not .NoMatch Then
                Do Until .EOF
                    If !Empresa <> g_empresa Or ![Codigo do Funcionario] <> x_funcionario Then
                        Exit Do
                    End If
                    If x_salario_familia = "S" Then
                        x_data_nascimento = ![Data de Nascimento]
                        If !Invalido Then
                            CalculaDependente = CalculaDependente + 1
                        ElseIf DateDiff("yyyy", x_data_nascimento, x_data) < 14 Then
                            CalculaDependente = CalculaDependente + 1
                        ElseIf DateDiff("yyyy", x_data_nascimento, x_data) = 14 And Not Month(x_data_nascimento) > Month(x_data) Then
                            CalculaDependente = CalculaDependente + 1
                        End If
                    Else
                        CalculaDependente = CalculaDependente + 1
                    End If
                    .MoveNext
                Loop
            End If
        End If
    End With
End Function
Function CalculaDiasTrabalhados() As Integer
    If IsDate(tbl_funcionario![Data de Admissao]) Then
        If Year(tbl_funcionario![Data de Admissao]) = Val(Mid(l_ano_mes, 1, 4)) And Month(tbl_funcionario![Data de Admissao]) = Val(Mid(l_ano_mes, 5, 2)) Then
            CalculaDiasTrabalhados = 31 - Day(tbl_funcionario![Data de Admissao])
            Exit Function
        End If
    End If
    CalculaDiasTrabalhados = Val(tbl_tabela_provento_desconto!Fracao)
End Function
Function CalculaPercentualINSS() As Currency
    Dim x_valor As Currency
    CalculaPercentualINSS = 0
    x_valor = CalculaBaseCalculoComposta
    If x_valor <= tbl_tabela_folha![Valor Final 1] Then
        CalculaPercentualINSS = tbl_tabela_folha![Percentual 1]
    ElseIf x_valor <= tbl_tabela_folha![Valor Final 2] Then
        CalculaPercentualINSS = tbl_tabela_folha![Percentual 2]
    ElseIf x_valor <= tbl_tabela_folha![Valor Final 3] Then
        CalculaPercentualINSS = tbl_tabela_folha![Percentual 3]
    ElseIf x_valor <= tbl_tabela_folha![Valor Final 4] Then
        CalculaPercentualINSS = tbl_tabela_folha![Percentual 4]
    End If
End Function
Function CalculaProventoDesconto(x_codigo As Integer, x_quantidade As Currency) As Currency
    Dim x_base_calculo
    CalculaProventoDesconto = 0
    x_base_calculo = CalculaBaseCalculo(x_codigo)
    'Salário Base
    If x_codigo = 1 Then
        If Mid(l_ano_mes, 5, 2) = "13" Then
            If x_quantidade = 12 Then
                CalculaProventoDesconto = x_base_calculo / 2
            Else
                CalculaProventoDesconto = x_base_calculo / 12 * x_quantidade / 2
            End If
        ElseIf Mid(l_ano_mes, 5, 2) = "14" Then
            If x_quantidade = 12 Then
                CalculaProventoDesconto = x_base_calculo
            Else
                CalculaProventoDesconto = x_base_calculo / 12 * x_quantidade
            End If
        Else
            If x_quantidade = 30 Then
                CalculaProventoDesconto = x_base_calculo
            Else
                CalculaProventoDesconto = x_base_calculo / 30 * x_quantidade
            End If
        End If
    'Adicional de Periculosidade
    ElseIf x_codigo = 5 Then
        CalculaProventoDesconto = x_base_calculo * x_quantidade / 100
    'Adicional Noturno
    ElseIf x_codigo = 10 Then
        CalculaProventoDesconto = x_base_calculo * x_quantidade / 100
    'Salário Família
    ElseIf x_codigo = 15 Then
        CalculaProventoDesconto = x_base_calculo * x_quantidade
    'Adiantamento de Salário
    ElseIf x_codigo = 500 Then
        CalculaProventoDesconto = x_base_calculo * x_quantidade / 100
    'Adiantamento 1a Parcela 13o Salário
    ElseIf x_codigo = 505 Then
        CalculaProventoDesconto = x_base_calculo
    'Faltas
    ElseIf x_codigo = 510 Then
        CalculaProventoDesconto = x_base_calculo / 30 * x_quantidade
    'D.S.R. Descontado
    ElseIf x_codigo = 515 Then
        CalculaProventoDesconto = x_base_calculo / 30 * x_quantidade
    'I.N.S.S
    ElseIf x_codigo = 520 Then
        CalculaProventoDesconto = x_base_calculo * x_quantidade / 100
    'Vale Transporte
    ElseIf x_codigo = 530 Then
        CalculaProventoDesconto = x_base_calculo * x_quantidade / 100
    'Seguro de Vida
    ElseIf x_codigo = 550 Then
        CalculaProventoDesconto = tbl_tabela_provento_desconto!valor * x_quantidade
    'Cesta Básica
    ElseIf x_codigo = 560 Then
        CalculaProventoDesconto = x_base_calculo * x_quantidade / 100
    'Falta de Caixa
    ElseIf x_codigo = 570 Then
        CalculaProventoDesconto = TotalizaFaltaCaixa
    'Contribuição Sindical
    ElseIf x_codigo = 590 Then
        CalculaProventoDesconto = x_base_calculo / 30 * x_quantidade
    'Contribuição Assistencial
    ElseIf x_codigo = 610 Then
        CalculaProventoDesconto = tbl_tabela_provento_desconto!valor * x_quantidade
    End If
    CalculaProventoDesconto = Format(CalculaProventoDesconto, "########0.00")
End Function
Private Sub CriaMovimentoAutomatico(x_funcionario As Integer, x_data As String)
    Dim x_quantidade As Currency
    Dim x_valor As Currency
    l_funcionario = x_funcionario
    With tbl_tabela_provento_desconto
        If .RecordCount > 0 Then
            .MoveFirst
            Do Until .EOF
                If !Automatico Then
                    If Val(Mid(l_ano_mes, 5, 2)) <= 12 Or (Val(Mid(l_ano_mes, 5, 2)) > 12 And !Codigo <= 5) Or (Val(Mid(l_ano_mes, 5, 2)) = 14 And (!Codigo = 505 Or !Codigo = 520)) Then
                        If AceitaProventoDesconto(!Codigo, x_funcionario) Then
                            x_quantidade = PreparaQuantidade(!Codigo, x_funcionario)
                            x_valor = Format(CalculaProventoDesconto(!Codigo, x_quantidade), "###,##0.00")
                            If x_valor > 0 Then
                                tbl_movimento_folha.AddNew
                                tbl_movimento_folha!Empresa = g_empresa
                                tbl_movimento_folha![Mes Ano] = x_data
                                tbl_movimento_folha![Codigo do Funcionario] = x_funcionario
                                tbl_movimento_folha![Codigo do Movimento] = !Codigo
                                tbl_movimento_folha!Quantidade = x_quantidade
                                tbl_movimento_folha!valor = x_valor
                                tbl_movimento_folha!Status = "Automatico"
                                tbl_movimento_folha.Update
                                l_ano_mes = x_data
                                l_funcionario = x_funcionario
                                l_codigo = !Codigo
                            End If
                        End If
                    End If
                End If
                .MoveNext
            Loop
        End If
    End With
End Sub
Private Sub DeletaMovimentoAutomatico(x_funcionario As Integer, x_data As String)
    With tbl_movimento_folha
        If .RecordCount > 0 Then
            .Index = "id_funcionario"
            .Seek ">=", g_empresa, x_funcionario, x_data, 0
            If Not .NoMatch Then
                Do Until .EOF
                    If !Empresa <> g_empresa Or ![Codigo do Funcionario] <> x_funcionario Or ![Mes Ano] <> x_data Then
                        Exit Do
                    End If
                    tbl_tabela_provento_desconto.Seek "=", ![Codigo do Movimento]
                    If Not tbl_tabela_provento_desconto.NoMatch Then
                        If tbl_tabela_provento_desconto!Automatico Then
                            .Edit
                            .Delete
                        End If
                    End If
                    .MoveNext
                Loop
            End If
            .Index = "id_data"
        End If
    End With
End Sub
Private Sub DesativaBotoes()
    cmd_novo.Enabled = False
    cmd_alterar.Enabled = False
    cmd_pesquisa.Enabled = False
    cmd_excluir.Enabled = False
    cmd_imprimir.Enabled = False
    cmd_sair.Enabled = False
    frm_move.Visible = False
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
End Sub
Private Sub Finaliza()
    tbl_dependente.Close
    tbl_empresa.Close
    tbl_funcionario.Close
    tbl_movimento_falta_caixa.Close
    tbl_movimento_falta_funcionario.Close
    tbl_movimento_folha.Close
    tbl_tabela_folha.Close
    tbl_tabela_provento_desconto.Close
End Sub
Private Sub PesquisaBaseCalculo()
    'MontaGridBaseCalculo
    'With tbl_tabela_provento_desconto
    '    .Seek ">=", l_codigo
    '    If Not .NoMatch Then
    '        Do Until .EOF
    '            AdcionaDadosGridBaseCalculo
    '            .MoveNext
    '        Loop
    '    End If
    '    Call BuscaRegistro(l_codigo)
    '    grid_provento_desconto.Row = grid_provento_desconto.Rows - 1
    '    grid_provento_desconto.Col = 0
    'End With
End Sub
Function PreparaQuantidade(x_codigo As Integer, x_funcionario As Integer) As Currency
    PreparaQuantidade = 0
    If x_codigo = 1 Then
        If Mid(l_ano_mes, 5, 2) = "13" Or Mid(l_ano_mes, 5, 2) = "14" Then
            PreparaQuantidade = Calcula13Salario
        Else
'            PreparaQuantidade = Format(Val(tbl_tabela_provento_desconto!Fracao), "###,##0.00")
            PreparaQuantidade = Format(CalculaDiasTrabalhados, "###,##0.00")
        End If
    'Adicional de Periculosidade
    ElseIf x_codigo = 5 Then
        PreparaQuantidade = Format(tbl_tabela_provento_desconto!Percentual, "###,##0.00")
    'Adicional Noturno
    ElseIf x_codigo = 10 Then
        PreparaQuantidade = Format(tbl_tabela_provento_desconto!Percentual, "###,##0.00")
    'Salário Família
    ElseIf x_codigo = 15 Then
        PreparaQuantidade = Format(CalculaDependente(x_funcionario, "S"), "###,##0.00")
    'Adiantamento de Salário
    ElseIf x_codigo = 500 Then
        PreparaQuantidade = Format(tbl_tabela_provento_desconto!Percentual, "###,##0.00")
    'Desconto Adiantamento 1a Parcela 13o Salário
    ElseIf x_codigo = 505 Then
        PreparaQuantidade = Format(tbl_tabela_provento_desconto!Percentual, "###,##0.00")
    'Faltas
    ElseIf x_codigo = 510 Then
        PreparaQuantidade = Format(QuantidadeFaltas(x_funcionario, "N"), "###,##0.00")
    'D.S.R. Remunerado
    ElseIf x_codigo = 515 Then
        PreparaQuantidade = Format(QuantidadeFaltas(x_funcionario, "S"), "###,##0.00")
    'I.N.S.S.
    ElseIf x_codigo = 520 Then
        PreparaQuantidade = Format(CalculaPercentualINSS, "###,##0.00")
    'Vale Transporte
    ElseIf x_codigo = 530 Then
        PreparaQuantidade = Format(tbl_tabela_provento_desconto!Percentual, "###,##0.00")
    'Seguro de Vida
    ElseIf x_codigo = 550 Then
        PreparaQuantidade = Format(1, "###,##0.00")
    'Cesta Basica
    ElseIf x_codigo = 560 Then
        If QuantidadeFaltas(x_funcionario, "N") = 0 Then
            PreparaQuantidade = Format(tbl_tabela_provento_desconto!valor, "###,##0.00")
        Else
            PreparaQuantidade = Format(tbl_tabela_provento_desconto!Percentual, "###,##0.00")
        End If
    'Contribuicao Sindical
    ElseIf x_codigo = 590 Then
        PreparaQuantidade = Format(1, "###,##0.00")
    'Contribuicao Assistencial
    ElseIf x_codigo = 610 And Not tbl_tabela_provento_desconto!Automatico = False Then
        PreparaQuantidade = Format(1, "###,##0.00")
    End If
End Function
Private Sub ProcessamentoAutomatico(x_funcionario As Integer, x_data As String)
    With tbl_funcionario
        If .RecordCount > 0 Then
            If x_funcionario = 0 Then
                .Seek ">=", g_empresa, x_funcionario
            Else
                .Seek "=", g_empresa, x_funcionario
            End If
            If Not .NoMatch Then
                Do Until .EOF
                    l_funcionario2 = !Codigo
                    If !Empresa <> g_empresa Then
                        Exit Do
                    End If
                    If !Situacao = "A" And ![Salario Base] > 0 Then
                        Call DeletaMovimentoAutomatico(!Codigo, x_data)
                        Call CriaMovimentoAutomatico(!Codigo, x_data)
                    End If
                    If x_funcionario > 0 Then
                        Exit Do
                    End If
                    .MoveNext
                Loop
            End If
            Call BuscaRegistro(l_ano_mes, l_funcionario, l_codigo)
        End If
    End With
End Sub
Private Sub cmd_alterar_Click()
    lOpcao = 2
    DesativaBotoes
    cmd_alterar.Visible = True
    cmd_alterar.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    frmDados.Enabled = True
    'tab_dados.Enabled = True
    txt_valor.SetFocus
End Sub
Private Sub cmd_anterior_Click()
    With tbl_movimento_folha
        If .RecordCount > 0 Then
            .MovePrevious
            If Not .BOF Then
                If !Empresa = g_empresa Then
                    AtualTela
                    Exit Sub
                End If
            End If
            MsgBox "Início de Arquivo.", 48, "Atenção!"
            .MoveNext
            cmd_proximo.SetFocus
        End If
    End With
End Sub
Private Sub cmd_cancelar_Click()
    LimpaTela
    If BuscaRegistro(l_ano_mes, l_funcionario, l_codigo) Then
        AtivaBotoes
        If cmd_alterar.Enabled Then
            cmd_alterar.SetFocus
        Else
            cmd_novo.SetFocus
        End If
    Else
        DesativaBotoes
        cmd_novo.Enabled = True
        cmd_sair.Enabled = True
        cmd_novo.SetFocus
    End If
    lOpcao = 0
End Sub
Sub LimpaGridProventoDesconto()
    Do Until grid_provento_desconto.Rows = 2
        grid_provento_desconto.Row = grid_provento_desconto.Rows - 1
        grid_provento_desconto.RemoveItem grid_provento_desconto.Row
    Loop
    grid_provento_desconto.Row = 1
    grid_provento_desconto.Col = 0
    grid_provento_desconto.Text = ""
    grid_provento_desconto.Col = 1
    grid_provento_desconto.Text = ""
    grid_provento_desconto.Col = 2
    grid_provento_desconto.Text = ""
    grid_provento_desconto.Col = 3
    grid_provento_desconto.Text = ""
End Sub
Private Sub LimpaTela()
    If l_gravados = 0 Then
        msk_data = "__/____"
        txt_funcionario = ""
        dbcbo_funcionario.BoundText = ""
    End If
    txt_provento_desconto = ""
    dbcbo_provento_desconto.BoundText = ""
    txt_quantidade = ""
    txt_valor = ""
    lbl_liquido = ""
End Sub
Private Sub cmd_excluir_Click()
    If IsDate("01/" & msk_data) Then
        If (MsgBox("Sim - Desejo excluir os proventos/descontos que estão no grid." & Chr(10) & "Não - Desejo excluir somente o registro provento/desconto.", vbYesNo + vbDefaultButton2, "Tipo de Exclusão") = vbYes) Then
            With tbl_movimento_folha
                .Seek ">=", g_empresa, l_ano_mes, l_funcionario, 0
                If Not .NoMatch Then
                    Do Until .EOF
                        If !Empresa <> g_empresa Or ![Mes Ano] <> l_ano_mes Or ![Codigo do Funcionario] <> l_funcionario Then
                            Exit Do
                        End If
                        .Edit
                        .Delete
                        .MoveNext
                    Loop
                    If Not BuscaDados Then
                        DesativaBotoes
                        cmd_novo.Enabled = True
                        cmd_sair.Enabled = True
                        cmd_novo.SetFocus
                    End If
                End If
            End With
        Else
            If (MsgBox("Deseja realmente excluir este registro?", 4 + 32 + 256, "Exclusão de Registro!")) = 6 Then
                lOpcao = 3
                tbl_movimento_folha.Edit
                tbl_movimento_folha.Delete
                If Not BuscaDados Then
                    DesativaBotoes
                    cmd_novo.Enabled = True
                    cmd_sair.Enabled = True
                    cmd_novo.SetFocus
                End If
                lOpcao = 0
            End If
        End If
    End If
End Sub
Private Sub cmd_imprimir_Click()
    If SelecionaImpressoraEpson(Me) Then
        SelecionaPaginaContraCheque
'        DesativaBotoes
        Relatorio
    End If
End Sub
Private Sub cmd_novo_Click()
    'Exit Sub
    'With tbl_movimento_folha
    '    .MoveFirst
    '    Do Until .EOF
    '        .Edit
    '        ![Mes Ano] = Format(!Data, "yyyymm")
    '        .Update
    '        .MoveNext
    '    Loop
    'End With
    'Exit Sub
    LimpaTela
    Inclui
    frmDados.Enabled = True
    'tab_dados.Enabled = True
    If l_gravados = 0 Then
        msk_data.SetFocus
    Else
        txt_provento_desconto.SetFocus
    End If
End Sub
Private Sub cmd_ok_Click()
    On Error GoTo FileError
    If ValidaCampos Then
        AtivaBotoes
        If lOpcao = 1 Then
            tbl_movimento_folha.AddNew
            AtualTabe
            tbl_movimento_folha.Update
            l_gravados = 1
        ElseIf lOpcao = 2 Then
            tbl_movimento_folha.Edit
            AtualTabe
            tbl_movimento_folha.Update
        End If
        Call BuscaRegistro(l_ano_mes, l_funcionario, l_codigo)
        lOpcao = 0
        cmd_novo.SetFocus
    End If
    Exit Sub
FileError:
    ErroArquivo tbl_movimento_folha.name, "Movimento da Folhao"
    Exit Sub
End Sub
Private Sub Relatorio()
    Dim posicao_y As Currency
    Dim tamanho_form As Integer
    Dim largura_form As Integer
    Dim x_linha As Currency
    Dim x_cgc As String
    Dim x_provento As Currency
    Dim x_desconto As Currency
    Dim x_tot_provento As Currency
    Dim x_tot_desconto As Currency
    Dim x_salario_base As Currency
    Dim x_base_inss As Currency
    Dim x_base_fgts As Currency
    Dim x_fgts As Currency
    Dim x_base_irrf As Currency
    Dim x_nome As String
    Dim x_data As Date
    Dim x_quantidade As String
    Dim x_imprime As Boolean
    Dim x_13_impresso As Boolean
    l_local = 1
    x_linha = -0.8
    x_tot_provento = 0
    x_tot_desconto = 0
        x_13_impresso = True
    If Val(Mid(l_ano_mes, 5, 2)) = 14 Then
        x_13_impresso = False
    End If
    'seleciona medidas para polegadas
    Printer.ScaleMode = 5
    'Seleciona largura do formulário
    Printer.ScaleWidth = 8
    largura_form = Printer.ScaleWidth
    'Seleciona altura do formulário
    Printer.ScaleHeight = 5.5
    tamanho_form = Printer.ScaleHeight
    'seleciona medidas para centímetros
    Printer.ScaleMode = 7
    'Seleciona nome da fonte
    Printer.FontName = "Arial"
    Printer.FontName = "Roman 10cpi"
    tbl_empresa.Seek "=", g_empresa
    If Not tbl_empresa.NoMatch Then
        x_cgc = Mid(tbl_empresa!CGC, 1, 2) & "."
        x_cgc = x_cgc & Mid(tbl_empresa!CGC, 3, 3) & "."
        x_cgc = x_cgc & Mid(tbl_empresa!CGC, 6, 3) & "/"
        x_cgc = x_cgc & Mid(tbl_empresa!CGC, 9, 4) & "-"
        x_cgc = x_cgc & Mid(tbl_empresa!CGC, 13, 2)
    End If
    Printer.FontName = "Roman 10cpi"
    Printer.FontBold = True
    ImprimeTexto "C.G.C.: " & x_cgc, 0, 9, x_linha + 0.8, l_local
    Printer.FontName = "Roman 6cpi"
    Printer.FontBold = True
    ImprimeTexto UCase(tbl_empresa!Nome), 0, 13, x_linha + 1.3, l_local
    Printer.FontBold = True
    Printer.FontName = "Roman 10cpi"
    If Val(Mid(l_ano_mes, 5, 2)) <= 12 Then
        x_data = "01/" & Mid(l_ano_mes, 5, 2) & "/" & Mid(l_ano_mes, 1, 4)
        ImprimeTexto Trim(Format(x_data, "mmmm")) & "/" & Format(x_data, "yyyy"), 13.5, 17, x_linha + 1.3, l_local
    ElseIf Val(Mid(l_ano_mes, 5, 2)) = 13 Then
        ImprimeTexto "Novembro/" & Mid(l_ano_mes, 1, 4), 13.5, 17, x_linha + 1.3, l_local
    ElseIf Val(Mid(l_ano_mes, 5, 2)) = 14 Then
        ImprimeTexto "Dezembro/" & Mid(l_ano_mes, 1, 4), 13.5, 17, x_linha + 1.3, l_local
    End If
    Printer.FontBold = False
    ImprimeTexto Format(dbcbo_funcionario.BoundText, "000"), 0, 2, x_linha + 2.4, l_local
    Printer.FontBold = True
    ImprimeTexto dbcbo_funcionario, 1.1, 12, x_linha + 2.4, l_local
    Printer.FontBold = False
    ImprimeTexto UCase(tbl_funcionario!Cargo), 1.1, 12, x_linha + 3, l_local
    x_linha = 2.9
    With tbl_movimento_folha
        If .RecordCount > 0 Then
            .Index = "id_funcionario"
            .Seek ">=", g_empresa, l_funcionario, l_ano_mes, 0
            Do Until .EOF
                If !Empresa <> g_empresa Or ![Codigo do Funcionario] <> l_funcionario Or ![Mes Ano] <> l_ano_mes Then
                    Exit Do
                End If
                x_provento = 0
                x_desconto = 0
                tbl_tabela_provento_desconto.Seek "=", ![Codigo do Movimento]
                If Not tbl_tabela_provento_desconto.NoMatch Then
                    x_nome = tbl_tabela_provento_desconto!Nome
                    If tbl_tabela_provento_desconto![Provento ou Desconto] = "P" Then
                        x_provento = !valor
                        x_tot_provento = x_tot_provento + !valor
                    Else
                        x_desconto = !valor
                        x_tot_desconto = x_tot_desconto + !valor
                    End If
                Else
                    x_nome = "** Nao Cadastrado **"
                End If
                x_quantidade = ""
                If ![Codigo do Movimento] = 1 Then
                    x_quantidade = Format(!Quantidade, "#0")
                ElseIf ![Codigo do Movimento] = 5 Then
                    x_quantidade = Format(!Quantidade, "#0") & "%"
                ElseIf ![Codigo do Movimento] = 10 Then
                    x_quantidade = Format(!Quantidade, "#0") & "%"
                ElseIf ![Codigo do Movimento] = 15 Then
                    x_quantidade = Format(!Quantidade, "#0")
                ElseIf ![Codigo do Movimento] = 500 Then
                    x_quantidade = Format(!Quantidade, "#0.00") & "%"
                ElseIf ![Codigo do Movimento] = 505 Then
                    x_quantidade = Format(!Quantidade, "#0") & "%"
                ElseIf ![Codigo do Movimento] = 520 Then
                    x_quantidade = Format(!Quantidade, "#0.00") & "%"
                ElseIf ![Codigo do Movimento] = 530 Then
                    x_quantidade = Format(!Quantidade, "#0") & "%"
                ElseIf ![Codigo do Movimento] = 560 Then
                    x_quantidade = Format(!Quantidade, "#0") & "%"
                End If
                x_imprime = True
                If Val(Mid(![Mes Ano], 5, 2)) = 13 Then
                    x_imprime = False
                ElseIf Val(Mid(![Mes Ano], 5, 2)) = 14 And ![Codigo do Movimento] < 500 Then
                    x_imprime = False
                End If
                'Imprime 13o Salário 2a Parcela
                If Not x_13_impresso And ![Codigo do Movimento] >= 500 Then
                    x_13_impresso = True
                    x_quantidade = ![Codigo do Movimento]
                    .Index = "id_data"
                    Call BuscaRegistro2(l_ano_mes, l_funcionario, 1)
                    x_linha = x_linha + 0.5
                    Printer.FontName = "Roman 12cpi"
                    ImprimeTexto "13o Salario", 1.1, 7.8, x_linha, l_local
                    ImprimeValor Format(!Quantidade, "#0") & "/12", 8.7, 10.2, x_linha, l_local
                    ImprimeValor Format(x_tot_provento, "##,###,##0.00"), 11, 13.8, x_linha, l_local
                    .Index = "id_funcionario"
                    .Seek ">=", g_empresa, l_funcionario, l_ano_mes, Val(x_quantidade)
                    x_quantidade = Format(!Quantidade, "#0") & "%"
                End If
                If x_imprime Then
                    x_linha = x_linha + 0.5
                    ImprimeTexto Format(![Codigo do Movimento], "000"), 0, 2, x_linha, l_local
                    Printer.FontName = "Roman 12cpi"
                    ImprimeTexto x_nome, 1.1, 7.8, x_linha, l_local
                    ImprimeValor x_quantidade, 8.7, 10.2, x_linha, l_local
                    Printer.FontName = "Roman 10cpi"
                    If x_provento > 0 Then
                        ImprimeValor Format(x_provento, "##,###,##0.00"), 11, 13.8, x_linha, l_local
                    Else
                        ImprimeValor Format(x_desconto, "##,###,##0.00"), 14.5, 17.4, x_linha, l_local
                    End If
                End If
                .MoveNext
            Loop
            If Val(Mid(l_ano_mes, 5, 2)) = 13 Then
                .Index = "id_data"
                Call BuscaRegistro2(l_ano_mes, l_funcionario, 1)
                x_linha = 2.9
                x_linha = x_linha + 0.5
                Printer.FontName = "Roman 12cpi"
                ImprimeTexto "13o Salario - 1a PARCELA", 1.1, 7.8, x_linha, l_local
                ImprimeValor Format(!Quantidade, "#0") & "/12", 8.7, 10.2, x_linha, l_local
                ImprimeValor Format(x_tot_provento, "##,###,##0.00"), 11, 13.8, x_linha, l_local
            End If
            .Index = "id_data"
            Call BuscaRegistro2(l_ano_mes, l_funcionario, l_codigo)
            x_linha = 10.3
            Printer.FontBold = True
            ImprimeValor Format(x_tot_provento, "##,###,##0.00"), 11, 13.8, x_linha, l_local
            ImprimeValor Format(x_tot_desconto, "##,###,##0.00"), 14.5, 17.4, x_linha, l_local
            x_tot_provento = x_tot_provento - x_tot_desconto
            x_linha = 11.1
            ImprimeValor Format(x_tot_provento, "##,###,##0.00"), 14.5, 17.4, x_linha, l_local
            Printer.FontBold = False
            x_linha = -0.8
            ImprimeCentralizado Trim(tbl_tabela_folha![Observacao 1]), 0, 10.2, x_linha + 10.8, l_local
            ImprimeCentralizado Trim(tbl_tabela_folha![Observacao 2]), 0, 10.2, x_linha + 11.3, l_local
            ImprimeCentralizado Trim(tbl_tabela_folha![Observacao 3]), 0, 10.2, x_linha + 11.8, l_local
        End If
    End With
    tbl_tabela_provento_desconto.Seek "=", 1
    x_salario_base = CalculaBaseCalculo(1)
    tbl_tabela_provento_desconto.Seek "=", 520
    x_base_inss = CalculaBaseCalculo(520)
    x_base_fgts = CalculaBaseCalculo(520)
    If Val(Mid(l_ano_mes, 5, 2)) = 14 Then
        x_base_fgts = x_base_fgts / 2
    End If
    x_fgts = Format(x_base_fgts * 8 / 100, "###,##0.00")
    x_base_irrf = 0
    x_linha = 11.9
    ImprimeValor Format(x_salario_base, "##,###,##0.00"), 0, 2.9, x_linha, l_local
    ImprimeValor Format(x_base_inss, "##,###,##0.00"), 3.4, 6.3, x_linha, l_local
    ImprimeValor Format(x_base_fgts, "##,###,##0.00"), 6.5, 9.4, x_linha, l_local
    ImprimeValor Format(x_fgts, "##,###,##0.00"), 9.2, 12.1, x_linha, l_local
    ImprimeValor Format(x_base_irrf, "##,###,##0.00"), 12.1, 15, x_linha, l_local
    Call BuscaRegistro2(l_ano_mes, l_funcionario, l_codigo)
    Printer.EndDoc
End Sub
Function RetiraString(x_string As String, numero As Integer) As Integer
    RetiraString = 0
    Dim x_index As Integer
    Dim x_inicio As Integer
    Dim x_numero As Integer
    x_inicio = 1
    x_numero = 1
    If Len(x_string) > 0 Then
        Do Until x_index > Len(x_string)
            x_index = x_index + 1
            If Mid(x_string, x_index, 1) = "@" Then
                If x_numero = numero Then
                    RetiraString = Mid(x_string, x_inicio, x_index - x_inicio)
                    Exit Function
                End If
                x_index = x_index + 2
                x_numero = x_numero + 1
                x_inicio = x_index + 1
            End If
        Loop
    End If
End Function
Private Sub SelecionaPaginaContraCheque()
Dim Impressora As Printer
Dim NumeroFormulario As Integer
    On Error GoTo PrinterError
    NumeroFormulario = 256
    Printer.PaperSize = NumeroFormulario
    Printer.ScaleMode = 5
'    For NumeroFormulario = 1 To 256
'        Printer.PaperSize = NumeroFormulario
'        'Seleciona largura do formulário
'        Printer.ScaleWidth = 8
'        'Seleciona altura do formulário
'        Printer.ScaleHeight = 5.5
'    Next
    Exit Sub
PrinterError:
    NumeroFormulario = NumeroFormulario + 1
    Resume



    NumeroFormulario = 256
    Resume
End Sub
Function TotalizaFaltaCaixa() As Currency
    Dim x_data_i As Date
    Dim x_data_f As Date
    Dim x_data As Date
    x_data = "01/" & Mid(l_ano_mes, 5, 2) & "/" & Mid(l_ano_mes, 1, 4)
    TotalizaFaltaCaixa = 0
    If Month(x_data) = 1 Then
        x_data_i = CDate("26/12/" & Year(x_data) - 1)
    Else
        x_data_i = CDate("26/" & Month(x_data) - 1 & "/" & Year(x_data))
    End If
    x_data_f = CDate("25/" & Month(x_data) & "/" & Year(x_data))
    With tbl_movimento_falta_caixa
        If .RecordCount > 0 Then
            .Seek ">=", g_empresa, l_funcionario, x_data_i, " "
            If Not .NoMatch Then
                Do Until .EOF
                    If !Empresa <> g_empresa Or ![Codigo do Funcionario] <> l_funcionario Or !Data > x_data_f Then
                        Exit Do
                    End If
                    TotalizaFaltaCaixa = TotalizaFaltaCaixa + !valor
                    .MoveNext
                Loop
            End If
        End If
    End With
End Function
Function ValidaCampos() As Boolean
    ValidaCampos = False
    If Val(Mid(msk_data, 1, 2)) < 1 Or Val(Mid(msk_data, 1, 2)) > 14 Then
        MsgBox "Informe o mês entre 01 a 14.", 64, "Atenção!"
        msk_data.SetFocus
    ElseIf Val(Mid(msk_data, 4, 4)) < 1998 Or Val(Mid(msk_data, 4, 4)) > 2500 Then
        MsgBox "Informe o ano entre 1998 a 2500.", 64, "Atenção!"
        msk_data.SetFocus
    ElseIf Not Val(dbcbo_funcionario.BoundText) > 0 Then
        MsgBox "Selecione o funcionário.", 64, "Atenção!"
        dbcbo_funcionario.SetFocus
    ElseIf Not Val(dbcbo_provento_desconto.BoundText) > 0 Then
        MsgBox "Selecione um provento/desconto.", 64, "Atenção!"
        dbcbo_provento_desconto.SetFocus
    ElseIf Not fValidaValor2(txt_quantidade) > 0 Then
        MsgBox "Informe a quantidade.", 64, "Atenção!"
        txt_quantidade.SetFocus
    ElseIf Not fValidaValor2(txt_valor) > 0 Then
        MsgBox "Informe o valor.", 64, "Atenção!"
        txt_valor.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_pesquisa_Click()
    consulta_movimento_folha.Show 1
    If Len(g_string) > 0 Then
        l_ano_mes = RetiraGString(1)
        l_funcionario = RetiraGString(2)
        l_codigo = RetiraGString(3)
        Call BuscaRegistro(l_ano_mes, l_funcionario, l_codigo)
    End If
End Sub
Private Sub cmd_primeiro_Click()
    With tbl_movimento_folha
        If .RecordCount > 0 Then
            .Seek ">", g_empresa, "190001", 0, 0
            If Not .NoMatch Then
                If !Empresa = g_empresa Then
                    AtualTela
                    cmd_proximo.SetFocus
                    Exit Sub
                End If
            End If
            MsgBox "Não há registro nesta empresa.", 64, "Erro de Verificação!"
        End If
    End With
End Sub
Private Sub cmd_processamento_Click()
    If Not IsDate("01/" & msk_data) Then
        MsgBox "Informe o mês/ano.", 64, "Atenção!"
        msk_data.SetFocus
    Else
        If Val(dbcbo_funcionario.BoundText) = 0 Then
            If (MsgBox("Na empresa " & g_nome_empresa & Chr(10) & "Será feito o movimento no mês/ano " & msk_data & "." & Chr(10) & "De todos os funcionários." & Chr(10) & Chr(10) & "Deseja realmente fazer este processamento?", vbYesNo + 256, "Processamento Automático da Folha!")) = 6 Then
                Call ProcessamentoAutomatico(Val(dbcbo_funcionario.BoundText), Mid(msk_data, 4, 4) & Mid(msk_data, 1, 2))
                MsgBox "A empresa " & g_nome_empresa & Chr(10) & "Está processado o movimento do mês/ano " & msk_data & "." & Chr(10) & "De todos os funcionários.", vbInformation, "Processamento Automático Concluido!"
                cmd_cancelar_Click
            End If
        Else
            If (MsgBox("Na empresa " & g_nome_empresa & Chr(10) & "Será feito o movimento no mês/ano " & msk_data & "." & Chr(10) & "Do funcionário " & Trim(dbcbo_funcionario) & "." & Chr(10) & Chr(10) & "Deseja realmente fazer este processamento?", vbYesNo + 256, "Processamento Automático da Folha!")) = 6 Then
                Call ProcessamentoAutomatico(Val(dbcbo_funcionario.BoundText), Mid(msk_data, 4, 4) & Mid(msk_data, 1, 2))
                MsgBox "A empresa " & g_nome_empresa & Chr(10) & "Está processado o movimento do mês/ano " & msk_data & "." & Chr(10) & "Do funcionário " & Trim(dbcbo_funcionario) & ".", vbInformation, "Processamento Automático Concluido!"
                cmd_cancelar_Click
            End If
        End If
    End If
End Sub
Private Sub cmd_proximo_Click()
    With tbl_movimento_folha
        If .RecordCount > 0 Then
            .MoveNext
            If Not .EOF Then
                If !Empresa = g_empresa Then
                    AtualTela
                    Exit Sub
                End If
            End If
            MsgBox "Fim de Arquivo.", 48, "Atenção!"
            .MovePrevious
            cmd_anterior.SetFocus
        End If
    End With
End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub cmd_ultimo_Click()
    With tbl_movimento_folha
        If .RecordCount > 0 Then
            .Seek "<", g_empresa, "250012", 9999, 9999
            If Not .NoMatch Then
                If !Empresa = g_empresa Then
                    AtualTela
                    cmd_anterior.SetFocus
                    Exit Sub
                End If
            End If
            MsgBox "Não há registro nesta empresa.", 64, "Erro de Verificação!"
        End If
    End With
End Sub
Private Sub dbcbo_funcionario_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        txt_provento_desconto.SetFocus
    End If
End Sub
Private Sub dbcbo_funcionario_LostFocus()
    If dbcbo_funcionario.BoundText <> "" And lOpcao > 0 Then
        txt_funcionario = dbcbo_funcionario.BoundText
        txt_funcionario_LostFocus
        txt_provento_desconto.SetFocus
    Else
        cmd_processamento.SetFocus
    End If
End Sub
Private Sub dbcbo_provento_desconto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        txt_quantidade.SetFocus
    End If
End Sub
Private Sub dbcbo_provento_desconto_LostFocus()
    Dim x_data As String
    If dbcbo_provento_desconto.BoundText <> "" And lOpcao > 0 Then
        txt_provento_desconto = dbcbo_provento_desconto.BoundText
        txt_provento_desconto_LostFocus
'        x_data = "01/" & msk_data
'        If lOpcao = 1 And IsDate(x_data) Then
'            tbl_tabela_folha.Seek "=", CDate(x_data)
'            If Not tbl_tabela_folha.NoMatch Then
'                MsgBox "Já existe tabela de premiação cadastrada nesta data." & Chr(10) & Chr(10) & "Mude a data informada.", 64, "Duplicidade de Registro!"
'                msk_data.SetFocus
'            End If
'        End If
        'Salário Base
        txt_quantidade = PreparaQuantidade(Val(dbcbo_provento_desconto.BoundText), Val(dbcbo_funcionario.BoundText))
        txt_quantidade_LostFocus
    Else
        cmd_processamento.SetFocus
    End If
End Sub
Private Sub Form_Activate()
    If g_empresa <> l_empresa Then
        flag_movimento_folha = 0
    End If
    If flag_movimento_folha = 0 Then
        dta_funcionario.RecordSource = "Select * From Funcionario Where Empresa = " & g_empresa & " And Situacao = " & Chr(34) & "A" & Chr(34) & " And [Salario Base] > " & 0 & " Order By [Nome]"
        dta_funcionario.Refresh
        dta_provento_desconto.RecordSource = "Select * From Tabela_Provento_Desconto Order By [Nome]"
        dta_provento_desconto.Refresh
        l_gravados = 0
        lOpcao = 0
        l_empresa = g_empresa
        DesativaBotoes
        If BuscaDados Then
            AtivaBotoes
        Else
            cmd_novo.Enabled = True
            cmd_sair.Enabled = True
        End If
        If cmd_novo.Enabled Then
            cmd_novo.SetFocus
        End If
    Else
        flag_movimento_folha = 0
    End If
    Screen.MousePointer = 1
End Sub
Private Sub Form_Deactivate()
    flag_movimento_folha = 1
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 And lOpcao = 0 Then
        KeyCode = 0
        cmd_novo_Click
    ElseIf KeyCode = vbKeyF3 And lOpcao = 0 Then
        KeyCode = 0
        cmd_alterar_Click
    ElseIf KeyCode = vbKeyF4 And Shift = 0 And lOpcao = 0 Then
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
    CentraForm Me
    Set tbl_dependente = bd_sgp.OpenTable("Dependente")
    Set tbl_empresa = bd_sgp.OpenTable("Empresas")
    Set tbl_funcionario = bd_sgp.OpenTable("Funcionario")
    Set tbl_movimento_falta_caixa = bd_sgp.OpenTable("Movimento_Falta_Caixa")
    Set tbl_movimento_falta_funcionario = bd_sgp.OpenRecordset("Movimento_Falta_Funcionario", dbOpenTable)
    Set tbl_movimento_folha = bd_sgp.OpenTable("Movimento_Folha")
    Set tbl_tabela_folha = bd_sgp.OpenTable("Tabela_Folha")
    Set tbl_tabela_provento_desconto = bd_sgp.OpenTable("Tabela_Provento_Desconto")
    tbl_dependente.Index = "id_codigo"
    tbl_empresa.Index = "id_codigo"
    tbl_funcionario.Index = "id_codigo"
    tbl_movimento_falta_caixa.Index = "id_funcionario"
    tbl_movimento_falta_funcionario.Index = "id_funcionario"
    tbl_movimento_folha.Index = "id_data"
    tbl_tabela_folha.Index = "id_mes_ano"
    tbl_tabela_provento_desconto.Index = "id_codigo"
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub MarcaCelulaProventoDesconto()
    grid_provento_desconto.Col = 0
    If grid_provento_desconto.Text <> "" Then
        grid_provento_desconto.Col = 0
        l_codigo = Val(grid_provento_desconto.Text)
        Call BuscaRegistro(l_ano_mes, l_funcionario, l_codigo)
        cmd_alterar.SetFocus
    End If
End Sub
Private Sub MontaGridProventoDesconto()
    Dim x_provento As Currency
    Dim x_desconto As Currency
    Dim x_nome As String
    Dim i As Integer
    lbl_liquido = ""
    LimpaGridProventoDesconto
    grid_provento_desconto.ForeColor = vbBlue
    grid_provento_desconto.Row = 0
    grid_provento_desconto.Col = 0
    grid_provento_desconto.Text = "Código"
    grid_provento_desconto.ColWidth(0) = 900
    grid_provento_desconto.ColAlignment(0) = 2
   'obs: o "9"equivale ao tab
    '0 = left, 1 = right ,2 =  center
    grid_provento_desconto.Col = 1
    grid_provento_desconto.Text = "Provento/Desconto"
    grid_provento_desconto.ColWidth(1) = 4100
    grid_provento_desconto.ColAlignment(1) = 0
    grid_provento_desconto.Col = 2
    grid_provento_desconto.Text = "Valor Provento"
    grid_provento_desconto.ColWidth(2) = 1080
    grid_provento_desconto.ColAlignment(2) = 1
    grid_provento_desconto.Col = 3
    grid_provento_desconto.Text = "Valor Provento"
    grid_provento_desconto.ColWidth(3) = 1080
    grid_provento_desconto.ColAlignment(3) = 1
    With tbl_movimento_folha
        If .RecordCount > 0 Then
            .Index = "id_funcionario"
            .Seek ">=", g_empresa, l_funcionario, Mid(msk_data, 4, 4) & Mid(msk_data, 1, 2), 0
            Do Until .EOF
                If !Empresa <> g_empresa Or ![Codigo do Funcionario] <> l_funcionario Or ![Mes Ano] <> Mid(msk_data, 4, 4) & Mid(msk_data, 1, 2) Then
                    Exit Do
                End If
                x_provento = 0
                x_desconto = 0
                tbl_tabela_provento_desconto.Seek "=", ![Codigo do Movimento]
                If Not tbl_tabela_provento_desconto.NoMatch Then
                    x_nome = tbl_tabela_provento_desconto!Nome
                    If tbl_tabela_provento_desconto![Provento ou Desconto] = "P" Then
                        x_provento = !valor
                    Else
                        x_desconto = !valor
                    End If
                Else
                    x_nome = "** Nao Cadastrado **"
                End If
                Call AdcionaDadosGridProventoDesconto(![Codigo do Movimento], x_nome, x_provento, x_desconto)
                .MoveNext
            Loop
            .Index = "id_data"
            Call BuscaRegistro2(l_ano_mes, l_funcionario, l_codigo)
            'Totaliza grid
            x_provento = 0
            x_desconto = 0
            For i = 1 To grid_provento_desconto.Rows - 1
                grid_provento_desconto.Row = i
                grid_provento_desconto.Col = 2
                x_provento = x_provento + fValidaValor2(grid_provento_desconto.Text)
                grid_provento_desconto.Col = 3
                x_desconto = x_desconto + fValidaValor2(grid_provento_desconto.Text)
            Next
            grid_provento_desconto.Row = grid_provento_desconto.Rows - 1
            grid_provento_desconto.Col = 1
            grid_provento_desconto.Text = "Total dos Proventos/Descontos"
            grid_provento_desconto.Col = 2
            grid_provento_desconto.Text = Format(x_provento, "###,##0.00") & "  "
            grid_provento_desconto.Col = 3
            grid_provento_desconto.Text = Format(x_desconto, "###,##0.00") & "  "
            grid_provento_desconto.SelStartRow = grid_provento_desconto.Rows - 1
            grid_provento_desconto.SelEndRow = grid_provento_desconto.Rows - 1
            grid_provento_desconto.SelStartCol = 0
            grid_provento_desconto.SelEndCol = 3
            lbl_liquido.ForeColor = vbBlue
            lbl_liquido = Format(x_provento - x_desconto, "###,##0.00") & "  "
            If grid_provento_desconto.Rows > 8 Then
                grid_provento_desconto.TopRow = grid_provento_desconto.Rows - 1 - 6
            End If
        End If
    End With
End Sub
Function QuantidadeFaltas(x_funcionario As Integer, x_justificada As String) As Integer
    Dim i As Integer
    Dim i2 As Integer
    Dim x_data_i As Date
    Dim x_data_f As Date
    Dim data_inicial(1 To 6) As Date
    Dim data_final(1 To 6) As Date
    Dim faltas(1 To 6) As Boolean
    Dim x_data As Date
    QuantidadeFaltas = 0
    'Calcula data inicial e final do período da folha
    x_data = "01/" & Mid(l_ano_mes, 5, 2) & "/" & Mid(l_ano_mes, 1, 4)
    If Month(x_data) = 1 Then
        x_data_i = CDate("26/12/" & Year(x_data) - 1)
    Else
        x_data_i = CDate("26/" & Month(x_data) - 1 & "/" & Year(x_data))
    End If
    x_data_f = CDate("25/" & Month(x_data) & "/" & Year(x_data))
    'Zera variáveis
    For i2 = 1 To 6
        data_inicial(i2) = "00:00:00"
        data_final(i2) = "00:00:00"
        faltas(i2) = False
    Next
    x_data = x_data_i
    data_inicial(1) = x_data
    x_data = x_data + (7 - Format(x_data, "w"))
    data_final(1) = x_data
    For i = 2 To 6
        x_data = x_data + 1
        If x_data > x_data_f Then
            Exit For
        End If
        data_inicial(i) = x_data
        i2 = 1
        Do Until i2 = 7
            x_data = x_data + 1
            If x_data > x_data_f Then
                x_data = x_data - 1
                Exit Do
            End If
            i2 = i2 + 1
        Loop
        data_final(i) = x_data
    Next
    With tbl_movimento_falta_funcionario
        If .RecordCount > 0 Then
            .Seek ">=", g_empresa, x_funcionario, x_data_i
            If Not .NoMatch Then
                Do Until .EOF
                    If !Empresa <> g_empresa Or ![Codigo do Funcionario] <> x_funcionario Or !Data > x_data_f Then
                        Exit Do
                    End If
                    If !Abonada = False Then
                        For i2 = 1 To 6
                            If !Data >= data_inicial(i2) And !Data <= data_final(i2) Then
                                faltas(i2) = True
                                Exit For
                            End If
                        Next
                        If !Justificada = False Then
                            QuantidadeFaltas = QuantidadeFaltas + 1
                        Else
                            If x_justificada = "S" Then
                                QuantidadeFaltas = QuantidadeFaltas + 1
                            End If
                        End If
                    End If
                    .MoveNext
                Loop
            End If
        End If
    End With
    If x_justificada = "S" Then
        QuantidadeFaltas = 0
        For i2 = 1 To 6
            If faltas(i2) = True Then
                QuantidadeFaltas = QuantidadeFaltas + 1
            End If
        Next
    End If
End Function
Function ExisteItemGridBaseCalculo(x_codigo As Integer) As Boolean
    ExisteItemGridBaseCalculo = False
    Dim i As Integer
    i = 1
    Do Until i >= grid_provento_desconto.Rows - 1
        grid_provento_desconto.Row = i
        grid_provento_desconto.Col = 0
        If Val(grid_provento_desconto.Text) = x_codigo Then
            ExisteItemGridBaseCalculo = True
            Exit Function
        End If
        i = i + 1
    Loop
End Function
Function ExisteRegistro() As Boolean
    ExisteRegistro = False
    With tbl_tabela_provento_desconto
        If .RecordCount > 0 Then
            .Index = "id_data"
            .Seek "=", l_codigo
            If Not .NoMatch Then
                MsgBox "Já existe movimento com este produto." & Chr(10) & Chr(10) & "Mude o produto informado.", 64, "Duplicidade de Registro!"
                ExisteRegistro = True
            End If
            .Index = "id_digitacao"
        End If
    End With
End Function
Private Sub grid_provento_desconto_DblClick()
    If lOpcao = 0 Then
        MarcaCelulaProventoDesconto
    End If
End Sub
Private Sub grid_provento_desconto_KeyDown(KeyCode As Integer, Shift As Integer)
    If lOpcao = 0 Then
        If KeyCode = vbKeyReturn Then
            MarcaCelulaProventoDesconto
        End If
    End If
End Sub
Private Sub msk_data_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_funcionario.SetFocus
    End If
End Sub
Private Sub msk_data_LostFocus()
    tbl_tabela_folha.Seek "=", Mid(msk_data, 4, 4) & Mid(msk_data, 1, 2)
    If tbl_tabela_folha.NoMatch Then
        MsgBox "Tabela da folha não cadastrada.", 64, "Erro de Consistência!"
        msk_data.SetFocus
    Else
        l_ano_mes = tbl_tabela_folha![Mes Ano]
    End If
End Sub
Private Sub txt_funcionario_GotFocus()
    If lOpcao = 1 Then
        txt_funcionario = ""
    End If
End Sub
Private Sub txt_funcionario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dbcbo_funcionario.SetFocus
    End If
End Sub
Private Sub txt_funcionario_LostFocus()
    If Val(txt_funcionario) > 0 And lOpcao > 0 Then
        tbl_funcionario.Seek "=", g_empresa, Val(txt_funcionario)
        If Not tbl_funcionario.NoMatch Then
            If tbl_funcionario!Situacao = "I" Then
                MsgBox "O funcionário " & Trim(tbl_funcionario!Nome) & " está inativo.", 64, "Atenção!"
                txt_funcionario.SetFocus
                Exit Sub
            Else
                dbcbo_funcionario.BoundText = tbl_funcionario!Codigo
                txt_provento_desconto.SetFocus
            End If
        Else
            MsgBox "Funcionário não cadastrado.", 64, "Atenção!"
            txt_funcionario.SetFocus
            Exit Sub
        End If
    End If
End Sub
Private Sub txt_provento_desconto_GotFocus()
    If lOpcao = 1 Then
        txt_provento_desconto = ""
    End If
End Sub
Private Sub txt_provento_desconto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dbcbo_provento_desconto.SetFocus
    End If
End Sub
Private Sub txt_provento_desconto_LostFocus()
    If Val(txt_provento_desconto) > 0 And lOpcao > 0 Then
        tbl_tabela_provento_desconto.Seek "=", Val(txt_provento_desconto)
        If Not tbl_tabela_provento_desconto.NoMatch Then
            dbcbo_provento_desconto.BoundText = tbl_tabela_provento_desconto!Codigo
            txt_quantidade.SetFocus
        Else
            MsgBox "Provento/Desconto não cadastrado.", 64, "Atenção!"
            txt_provento_desconto.SetFocus
            Exit Sub
        End If
    End If
End Sub
Private Sub txt_quantidade_GotFocus()
    txt_quantidade.SelStart = 0
    txt_quantidade.SelLength = Len(txt_quantidade)
End Sub
Private Sub txt_quantidade_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_valor.SetFocus
    End If
End Sub
Private Sub txt_quantidade_LostFocus()
    txt_quantidade = Format(txt_quantidade, "###,##0.00")
    txt_valor = Format(CalculaProventoDesconto(Val(txt_provento_desconto), fValidaValor2(txt_quantidade)), "###,##0.00")
    If fValidaValor2(txt_valor) > 0 Then
        cmd_ok.SetFocus
    End If
End Sub
Private Sub txt_valor_GotFocus()
    txt_valor.SelStart = 0
    txt_valor.SelLength = Len(txt_valor)
End Sub
Private Sub txt_valor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub txt_valor_LostFocus()
    txt_valor = Format(txt_valor, "###,##0.00")
End Sub
