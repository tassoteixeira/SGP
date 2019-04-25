VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frm_emissao_cheques_folhas 
   Caption         =   "Emissão de Cheques"
   ClientHeight    =   5520
   ClientLeft      =   1485
   ClientTop       =   2385
   ClientWidth     =   9015
   Icon            =   "Rel_chfo.frx":0000
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   Picture         =   "Rel_chfo.frx":030A
   ScaleHeight     =   5520
   ScaleWidth      =   9015
   Begin VB.CommandButton btnCopia 
      Caption         =   "&Cópia"
      Height          =   855
      Left            =   4560
      Picture         =   "Rel_chfo.frx":0750
      Style           =   1  'Graphical
      TabIndex        =   47
      ToolTipText     =   "Imprime cópia de cheque."
      Top             =   4560
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   5460
      Picture         =   "Rel_chfo.frx":1A2A
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   4560
      Width           =   795
   End
   Begin VB.CommandButton cmd_pesquisa 
      Caption         =   "&Pesquisa"
      Height          =   855
      Left            =   2760
      Picture         =   "Rel_chfo.frx":2D04
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Pesquisa um registro específico."
      Top             =   4560
      Width           =   795
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   1860
      Picture         =   "Rel_chfo.frx":3FDE
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Exclui o registro atual."
      Top             =   4560
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   960
      Picture         =   "Rel_chfo.frx":52B8
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Altera o registro atual."
      Top             =   4560
      Width           =   795
   End
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   60
      Picture         =   "Rel_chfo.frx":6592
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Cria um novo registro."
      Top             =   4560
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3660
      Picture         =   "Rel_chfo.frx":786C
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "Imprime cheque."
      Top             =   4560
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Enabled         =   0   'False
      Height          =   4395
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   8895
      Begin MSAdodcLib.Adodc adodcFuncionario 
         Height          =   330
         Left            =   3540
         Top             =   2100
         Visible         =   0   'False
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "adodcFuncionario"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.ComboBox cboChequePosse 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   3900
         Width           =   2175
      End
      Begin VB.Frame Frame1 
         Caption         =   "&Tipo de Cheque"
         Height          =   615
         Left            =   4080
         TabIndex        =   8
         Top             =   720
         Width           =   4515
         Begin VB.OptionButton optFormularioContinuo 
            Caption         =   "Formulário Contínuo"
            Height          =   255
            Left            =   180
            TabIndex        =   9
            Top             =   240
            Value           =   -1  'True
            Width           =   1875
         End
         Begin VB.OptionButton optChequeFolha 
            Caption         =   "Folha Avulsa"
            Height          =   255
            Left            =   2940
            TabIndex        =   10
            Top             =   240
            Width           =   1395
         End
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2880
         Picture         =   "Rel_chfo.frx":8B46
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_v 
         Height          =   315
         Left            =   2880
         Picture         =   "Rel_chfo.frx":9E20
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   2820
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_b 
         Height          =   315
         Left            =   2880
         Picture         =   "Rel_chfo.frx":B0FA
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   3540
         Width           =   495
      End
      Begin VB.CommandButton cmd_funcionario 
         Caption         =   "F&uncionários"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6660
         TabIndex        =   19
         ToolTipText     =   "Preenche o campo ""favorecido"" com funcionários."
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CommandButton cmd_fornecedores 
         Caption         =   "Fo&rnecedores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6660
         TabIndex        =   18
         ToolTipText     =   "Preenche o campo ""favorecido"" com fornecedores."
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox txt_nome 
         Height          =   300
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   17
         Top             =   2100
         Width           =   4935
      End
      Begin VB.TextBox txt_historico 
         Height          =   300
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   21
         Top             =   2460
         Width           =   4935
      End
      Begin VB.TextBox txt_numero 
         Height          =   300
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   4
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox msk_valor 
         Height          =   300
         Left            =   1680
         TabIndex        =   12
         Top             =   1380
         Width           =   1095
      End
      Begin MSMask.MaskEdBox msk_data_baixa 
         Height          =   315
         Left            =   1680
         TabIndex        =   30
         Top             =   3540
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_data_vencimento 
         Height          =   315
         Left            =   1680
         TabIndex        =   23
         Top             =   2820
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.OptionButton optSituacao 
         Caption         =   "Aberto"
         Height          =   195
         Index           =   0
         Left            =   1680
         TabIndex        =   26
         Top             =   3240
         Width           =   1095
      End
      Begin VB.OptionButton optSituacao 
         Caption         =   "Baixado"
         Height          =   195
         Index           =   1
         Left            =   2820
         TabIndex        =   27
         Top             =   3240
         Width           =   1095
      End
      Begin VB.OptionButton optSituacao 
         Caption         =   "Cancelado"
         Height          =   195
         Index           =   2
         Left            =   3960
         TabIndex        =   28
         Top             =   3240
         Width           =   1095
      End
      Begin MSMask.MaskEdBox msk_data 
         Height          =   315
         Left            =   1680
         TabIndex        =   6
         Top             =   960
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSAdodcLib.Adodc adodcConta 
         Height          =   330
         Left            =   3000
         Top             =   240
         Visible         =   0   'False
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "adodcConta"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataListLib.DataCombo dtcboConta 
         Bindings        =   "Rel_chfo.frx":C3D4
         Height          =   315
         Left            =   1680
         TabIndex        =   2
         Top             =   240
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   "dtcboConta"
      End
      Begin MSDataListLib.DataCombo dtcboFuncionario 
         Bindings        =   "Rel_chfo.frx":C3ED
         Height          =   315
         Left            =   1680
         TabIndex        =   16
         Top             =   2100
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   "dtcboFuncionario"
      End
      Begin VB.Label Label11 
         Caption         =   "Cheque em posse"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   3900
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "&Número da Conta"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Data da &Baixa"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   3600
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "&Data do Vencimento"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "S&ituação"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   25
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label lbl_extenso 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1680
         TabIndex        =   14
         Top             =   1740
         Width           =   4935
      End
      Begin VB.Label Label2 
         Caption         =   "&Histórico"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   2460
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Extenso"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1740
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "&Favorecido"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2100
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "&Número do Cheque"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Data de &Emissão"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1020
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "&Valor do Cheque"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1380
         Width           =   1455
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   60
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   6780
      TabIndex        =   42
      Top             =   4440
      Width           =   2175
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "Rel_chfo.frx":C40C
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "Rel_chfo.frx":D6E6
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "Rel_chfo.frx":E9C0
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "Rel_chfo.frx":FC9A
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   8160
      Picture         =   "Rel_chfo.frx":10F74
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Cancela o registro atual."
      Top             =   4560
      Width           =   795
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   7260
      Picture         =   "Rel_chfo.frx":1224E
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Confirma o registro atual."
      Top             =   4560
      Width           =   795
   End
End
Attribute VB_Name = "frm_emissao_cheques_folhas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flag_emissao_cheques_folhas As Integer

Private ChequeFolha As New cChequeFolha
Private ConfiguracaoCheque As New cConfiguracaoCheque

Dim lOpcao As Integer
Dim l_data As Date
Dim l_numero As String
Dim lCruzarCheque As Boolean
Private Sub AtivaBotoes()
    cmd_novo.Enabled = True
    cmd_alterar.Enabled = True
    cmd_excluir.Enabled = True
    cmd_imprimir.Enabled = True
    btnCopia.Enabled = True
    cmd_pesquisa.Enabled = True
    cmd_sair.Enabled = True
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
    frm_move.Visible = True
End Sub
Private Sub AtualTabe()
    ChequeFolha.Empresa = g_empresa
    ChequeFolha.numero = txt_numero.Text
    ChequeFolha.Data = Format(msk_data.Text, "dd/MM/yyyy")
    ChequeFolha.valor = fValidaValor2(msk_valor.Text)
    ChequeFolha.Nome = "" & txt_nome.Text
    ChequeFolha.Historico = "" & txt_historico.Text
    If optSituacao(0) Then
        ChequeFolha.Situacao = "A"
    ElseIf optSituacao(1) Then
        ChequeFolha.Situacao = "B"
    ElseIf optSituacao(2) Then
        ChequeFolha.Situacao = "C"
    End If
    If IsDate(msk_data_vencimento.Text) Then
        ChequeFolha.DataVencimento = Format(msk_data_vencimento.Text, "dd/MM/yyyy")
    Else
        ChequeFolha.DataVencimento = "00:00:00"
    End If
    If IsDate(msk_data_baixa.Text) Then
        ChequeFolha.DataBaixa = Format(msk_data_baixa.Text, "dd/MM/yyyy")
    Else
        ChequeFolha.DataBaixa = "00:00:00"
    End If
    If optChequeFolha.Value = True Then
        ChequeFolha.TipoCheque = 1
    Else
        ChequeFolha.TipoCheque = 2
    End If
    ChequeFolha.NumeroConta = dtcboConta.BoundText
    ChequeFolha.ChequeemPosse = cboChequePosse.ListIndex + 1
End Sub
Private Sub AtualTela()
    l_numero = ChequeFolha.numero
    l_data = ChequeFolha.Data
    dtcboConta.BoundText = ChequeFolha.NumeroConta
    txt_numero.Text = ChequeFolha.numero
    msk_data.Text = Format(ChequeFolha.Data, "dd/mm/yyyy")
    msk_valor.Text = Format(ChequeFolha.valor, "###,##0.00")
    If ChequeFolha.TipoCheque = 1 Then
        optChequeFolha.Value = True
    Else
        optFormularioContinuo.Value = True
    End If
    If fValidaValor2(msk_valor.Text) > 0 Then
        lbl_extenso.Caption = FazExtenso(msk_valor.Text)
    Else
        lbl_extenso.Caption = ""
    End If
    txt_nome.Text = ChequeFolha.Nome
    txt_historico.Text = ChequeFolha.Historico
    If ChequeFolha.Situacao = "A" Then
        optSituacao(0).Value = True
    ElseIf ChequeFolha.Situacao = "B" Then
        optSituacao(1).Value = True
    ElseIf ChequeFolha.Situacao = "C" Then
        optSituacao(2).Value = True
    End If
    If IsDate(ChequeFolha.DataVencimento) Then
        msk_data_vencimento.Text = Format(ChequeFolha.DataVencimento, "dd/mm/yyyy")
    Else
        msk_data_vencimento.Text = "__/__/____"
    End If
    If IsDate(ChequeFolha.DataBaixa) Then
        msk_data_baixa.Text = Format(ChequeFolha.DataBaixa, "dd/mm/yyyy")
    Else
        msk_data_baixa.Text = "__/__/____"
    End If
    cboChequePosse.ListIndex = ChequeFolha.ChequeemPosse - 1
    frm_dados.Enabled = False
End Sub
Private Sub DesativaBotoes()
    cmd_novo.Enabled = False
    cmd_alterar.Enabled = False
    cmd_excluir.Enabled = False
    cmd_imprimir.Enabled = False
    btnCopia.Enabled = False
    cmd_pesquisa.Enabled = False
    cmd_sair.Enabled = False
    frm_move.Visible = False
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
End Sub
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set ChequeFolha = Nothing
    Set ConfiguracaoCheque = Nothing
End Sub
Private Sub Inclui()
    lOpcao = 1
    DesativaBotoes
    cmd_novo.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
End Sub
Private Sub RelatorioChequeFolha(ByVal pCopiaCH As Boolean)
    Dim posicao_y As Currency
    Dim tamanho_form As Integer
    Dim largura_form As Integer
    Dim x_tamanho_string As Currency
    Dim x_extenso As String
    Dim x_asterisco As String
    'seleciona medidas para centímetros
    Printer.ScaleMode = 7
    'Seleciona Formulário de cheque
'@@@    'Printer.PaperSize = 52
    'Seleciona largura do formulário
    largura_form = Printer.ScaleWidth
    'Seleciona altura do formulário
    tamanho_form = Printer.ScaleHeight
    'Seleciona nome da fonte
    Printer.FontName = "Arial"
    Printer.FontName = "Times New Roman"
    Printer.CurrentX = 1.7
    Printer.CurrentY = 1
    Printer.Print " "
    ''Printer.FontBold = True
    Printer.FontName = "Times New Roman"
    Printer.FontSize = 12
    x_tamanho_string = Printer.TextWidth(lbl_extenso)
    If x_tamanho_string > 14.3 Then
        Printer.FontSize = 10
        x_tamanho_string = Printer.TextWidth(lbl_extenso)
        If x_tamanho_string > 14.3 Then
            MsgBox "Tamanho da String = " & x_tamanho_string, vbInformation, "Extenso muito longo!"
            Printer.EndDoc
            Exit Sub
        End If
    End If
    x_extenso = "( " & lbl_extenso
    Do Until x_tamanho_string >= 14.3
        x_extenso = x_extenso & "*"
        x_tamanho_string = Printer.TextWidth(x_extenso)
    Loop
'    Printer.CurrentX = 1.7
'    Printer.CurrentY = 1.4 + 0.2
    Printer.CurrentX = ConfiguracaoCheque.Extenso1Esquerda
    Printer.CurrentY = ConfiguracaoCheque.Extenso1Superior
    Printer.Print x_extenso
    Printer.FontSize = 16
'    Printer.CurrentX = 13 + 0.5
'    Printer.CurrentY = 0.4 + 0.2
    Printer.CurrentX = ConfiguracaoCheque.ValorEsquerda
    Printer.CurrentY = ConfiguracaoCheque.ValorSuperior
    Printer.Print Format(fValidaValor2(msk_valor.Text), "###,##0.00")
    Printer.Print
    If pCopiaCH = True Then
        Printer.CurrentX = 0
        Printer.CurrentY = ConfiguracaoCheque.ValorSuperior
        Printer.Print "C/C: " & dtcboConta.BoundText
        Printer.Print
        Printer.CurrentX = ConfiguracaoCheque.ValorEsquerda - 8
        Printer.CurrentY = ConfiguracaoCheque.ValorSuperior
        Printer.Print "Ch. N. " & txt_numero.Text
        Printer.Print
    End If
    Printer.FontSize = 12
'    Printer.CurrentX = 0.5
'    Printer.CurrentY = 2.3
    Printer.CurrentX = ConfiguracaoCheque.Extenso2Esquerda
    Printer.CurrentY = ConfiguracaoCheque.Extenso2Superior
    If txt_historico = "teste tasso" Then
        Printer.CurrentY = Printer.CurrentY - 0.2
    End If
    x_tamanho_string = 0
    x_asterisco = ""
    Do Until x_tamanho_string >= 15.5
        x_asterisco = x_asterisco & "*"
        x_tamanho_string = Printer.TextWidth(x_asterisco)
    Loop
    Printer.Print x_asterisco
    Printer.FontSize = 14
'    Printer.CurrentX = 0.5
'    Printer.CurrentY = 2.9
    Printer.CurrentX = ConfiguracaoCheque.FavorecidoEsquerda
    Printer.CurrentY = ConfiguracaoCheque.FavorecidoSuperior
    If txt_historico = "teste tasso" Then
        Printer.CurrentY = Printer.CurrentY - 0.3
    End If
    Printer.Print txt_nome
'    Printer.CurrentX = 8.5
'    Printer.CurrentY = 3.7 - 0.2
    Printer.CurrentX = ConfiguracaoCheque.CidadeEsquerda
    Printer.CurrentY = ConfiguracaoCheque.CidadeSuperior
    If txt_historico = "teste tasso" Then
        Printer.CurrentY = Printer.CurrentY - 0.6
    End If
    Printer.Print "Goiânia"
'    Printer.CurrentX = 10.5
'    Printer.CurrentY = 3.7 - 0.2
    Printer.CurrentX = ConfiguracaoCheque.DiaEsquerda
    Printer.CurrentY = ConfiguracaoCheque.DiaSuperior
    If txt_historico = "teste tasso" Then
        Printer.CurrentX = Printer.CurrentX + 0.2
        Printer.CurrentY = Printer.CurrentY - 0.6
    End If
    Printer.Print Day(msk_data.Text)
'    Printer.CurrentX = 12 - 0.3
'    Printer.CurrentY = 3.7 - 0.2
    Printer.CurrentX = ConfiguracaoCheque.MesEsquerda
    Printer.CurrentY = ConfiguracaoCheque.MesSuperior
    If txt_historico = "teste tasso" Then
        Printer.CurrentX = Printer.CurrentX + 0.2
        Printer.CurrentY = Printer.CurrentY - 0.6
    End If
    Printer.Print Format(msk_data.Text, "mmmm")
'    Printer.CurrentX = 15.7 + 0.2
'    Printer.CurrentY = 3.7 - 0.2
    Printer.CurrentX = ConfiguracaoCheque.AnoEsquerda
    Printer.CurrentY = ConfiguracaoCheque.AnoSuperior
    If txt_historico = "teste tasso" Then
        Printer.CurrentX = Printer.CurrentX + 0.2
        Printer.CurrentY = Printer.CurrentY - 0.6
    End If
    Printer.Print Mid(Year(msk_data.Text), 3, 2)
    
    
    If pCopiaCH = True Then
        Printer.CurrentX = ConfiguracaoCheque.CidadeEsquerda + 2
        Printer.CurrentY = ConfiguracaoCheque.CidadeSuperior + 1
        Printer.Print g_nome_empresa
        Printer.Print
        Printer.CurrentX = 0
        Printer.CurrentY = ConfiguracaoCheque.CidadeSuperior + 1.5
        Printer.Print dtcboConta.Text
        Printer.Print
    End If
    
    
    If IsDate(msk_data_vencimento.Text) Then
        Printer.FontSize = 16
        Printer.CurrentX = 11.5
        Printer.CurrentY = 5.8
        Printer.Print "BOM P/ " & Format(msk_data_vencimento.Text, "dd/mm/yyyy")
    End If
    
    If lCruzarCheque Then
        Printer.Line (9.5, 0)-(4, 6), RGB(0, 0, 0)
        Printer.Line (10.5, 0)-(5, 6), RGB(0, 0, 0)
    End If
    
    Printer.EndDoc
End Sub
Private Sub RelatorioChequeFormularioContinuo(ByVal pCopiaCH As Boolean)
    Dim posicao_y As Currency
    Dim tamanho_form As Integer
    Dim largura_form As Integer
    Dim x_tamanho_string As Currency
    Dim x_extenso As String
    Dim x_asterisco As String
    'seleciona medidas para centímetros
    Printer.ScaleMode = 7
    'Seleciona Formulário de cheque
'@@@    'Printer.PaperSize = 52
    'Seleciona largura do formulário
    largura_form = Printer.ScaleWidth
    'Seleciona altura do formulário
    tamanho_form = Printer.ScaleHeight
    'Seleciona nome da fonte
    Printer.FontName = "Arial"
    Printer.FontName = "Times New Roman"
    Printer.CurrentX = 1.7
    Printer.CurrentY = 1
    Printer.Print " "
    ''Printer.FontBold = True
    Printer.FontName = "Times New Roman"
    Printer.FontSize = 12
    x_tamanho_string = Printer.TextWidth(lbl_extenso)
    If x_tamanho_string > 14.3 Then
        Printer.FontSize = 10
        x_tamanho_string = Printer.TextWidth(lbl_extenso)
        If x_tamanho_string > 14.3 Then
            MsgBox "Tamanho da String = " & x_tamanho_string, vbInformation, "Extenso muito longo!"
            Printer.EndDoc
            Exit Sub
        End If
    End If
    x_extenso = "( " & lbl_extenso
    Do Until x_tamanho_string >= 14.3
        x_extenso = x_extenso & "*"
        x_tamanho_string = Printer.TextWidth(x_extenso)
    Loop
'    Printer.CurrentX = 1.7
'    Printer.CurrentY = 1.4 + 0.2
    Printer.CurrentX = ConfiguracaoCheque.Extenso1Esquerda
    Printer.CurrentY = ConfiguracaoCheque.Extenso1Superior
    Printer.Print x_extenso
    Printer.FontSize = 16
'    Printer.CurrentX = 13 + 0.5
'    Printer.CurrentY = 0.4 + 0.2
    Printer.CurrentX = ConfiguracaoCheque.ValorEsquerda
    Printer.CurrentY = ConfiguracaoCheque.ValorSuperior
    Printer.Print Format(fValidaValor2(msk_valor.Text), "###,##0.00")
    Printer.Print
    If pCopiaCH = True Then
        Printer.CurrentX = 0
        Printer.CurrentY = ConfiguracaoCheque.ValorSuperior
        Printer.Print "C/C: " & dtcboConta.BoundText
        Printer.Print
        Printer.CurrentX = ConfiguracaoCheque.ValorEsquerda - 8
        Printer.CurrentY = ConfiguracaoCheque.ValorSuperior
        Printer.Print "Ch. N. " & txt_numero.Text
        Printer.Print
    End If
    Printer.FontSize = 12
'    Printer.CurrentX = 0.5
'    Printer.CurrentY = 2.3
    Printer.CurrentX = ConfiguracaoCheque.Extenso2Esquerda
    Printer.CurrentY = ConfiguracaoCheque.Extenso2Superior
    If txt_historico = "teste tasso" Then
        Printer.CurrentY = Printer.CurrentY - 0.2
    End If
    x_tamanho_string = 0
    x_asterisco = ""
    Do Until x_tamanho_string >= 15.5
        x_asterisco = x_asterisco & "*"
        x_tamanho_string = Printer.TextWidth(x_asterisco)
    Loop
    Printer.Print x_asterisco
    Printer.FontSize = 14
'    Printer.CurrentX = 0.5
'    Printer.CurrentY = 2.9
    Printer.CurrentX = ConfiguracaoCheque.FavorecidoEsquerda
    Printer.CurrentY = ConfiguracaoCheque.FavorecidoSuperior
    Printer.Print txt_nome
'    Printer.CurrentX = 8.5
'    Printer.CurrentY = 3.7 - 0.2
    Printer.CurrentX = ConfiguracaoCheque.CidadeEsquerda
    Printer.CurrentY = ConfiguracaoCheque.CidadeSuperior
    Printer.Print "Goiânia"
'    Printer.CurrentX = 10.5
'    Printer.CurrentY = 3.7 - 0.2
    Printer.CurrentX = ConfiguracaoCheque.DiaEsquerda
    Printer.CurrentY = ConfiguracaoCheque.DiaSuperior
    Printer.Print Day(msk_data.Text)
'    Printer.CurrentX = 12 - 0.3
'    Printer.CurrentY = 3.7 - 0.2
    Printer.CurrentX = ConfiguracaoCheque.MesEsquerda
    Printer.CurrentY = ConfiguracaoCheque.MesSuperior
    Printer.Print Format(msk_data.Text, "mmmm")
'    Printer.CurrentX = 15.7 + 0.2
'    Printer.CurrentY = 3.7 - 0.2
    Printer.CurrentX = ConfiguracaoCheque.AnoEsquerda
    Printer.CurrentY = ConfiguracaoCheque.AnoSuperior
    Printer.Print Mid(Year(msk_data.Text), 3, 2)
    
    
    If pCopiaCH = True Then
        Printer.CurrentX = ConfiguracaoCheque.CidadeEsquerda + 2
        Printer.CurrentY = ConfiguracaoCheque.CidadeSuperior + 1
        Printer.Print g_nome_empresa
        Printer.Print
        Printer.CurrentX = 0
        Printer.CurrentY = ConfiguracaoCheque.CidadeSuperior + 1.5
        Printer.Print dtcboConta.Text
        Printer.Print
    End If
    
    
    If IsDate(msk_data_vencimento.Text) Then
        Printer.FontSize = 16
        Printer.CurrentX = 11.5
        Printer.CurrentY = 5.8
        Printer.Print "BOM P/ " & Format(msk_data_vencimento.Text, "dd/mm/yyyy")
    End If
    
    If lCruzarCheque Then
        Printer.Line (9.5, 0)-(4, 6), RGB(0, 0, 0)
        Printer.Line (10.5, 0)-(5, 6), RGB(0, 0, 0)
    End If
    
    Printer.EndDoc
End Sub
Private Sub btnCopia_Click()
    lCruzarCheque = False
    If MsgBox("Deseja Cruzar este cheque?", vbQuestion + vbYesNo + vbDefaultButton2, "Cruzar Cheque") = vbYes Then
        lCruzarCheque = True
    End If
    If Me.optChequeFolha Then
        If ConfiguracaoCheque.LocalizarCodigo(g_empresa, 1) Then
            If SelecionaImpressoraHP(Me) Then
                RelatorioChequeFolha (True)
                cmd_novo.SetFocus
            End If
        Else
            MsgBox "Não existe configuração para impressão de cheque.", vbInformation, "Ajustar Configuração de Cheque!"
        End If
    Else
        If ConfiguracaoCheque.LocalizarCodigo(g_empresa, 1) Then
            If SelecionaImpressoraEpson(Me) Then
                RelatorioChequeFormularioContinuo (True)
                cmd_novo.SetFocus
            End If
        Else
            MsgBox "Não existe configuração para impressão de cheque.", vbInformation, "Ajustar Configuração de Cheque!"
        End If
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
    txt_nome.SetFocus
End Sub
Private Sub cmd_anterior_Click()
    If ChequeFolha.LocalizarAnterior() Then
        AtualTela
    Else
        MsgBox "Início de Arquivo.", vbInformation, "Atenção!"
        cmd_proximo.SetFocus
    End If
End Sub
Private Sub cmd_cancelar_Click()
    LimpaTela
    If ChequeFolha.LocalizarUltimo(g_empresa) Then
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
    dtcboConta.BoundText = ""
    txt_numero.Text = ""
    msk_data.Text = "__/__/____"
    'optFormularioContinuo.Value = True
    msk_valor.Text = ""
    lbl_extenso.Caption = ""
    txt_nome.Text = ""
    txt_historico.Text = ""
    optSituacao(0).Value = True
    msk_data_vencimento.Text = "__/__/____"
    msk_data_baixa.Text = "__/__/____"
    cboChequePosse.ListIndex = -1
End Sub
Private Sub PreencheCboChequePosse()
    cboChequePosse.Clear
    cboChequePosse.AddItem "1 Favorecido"
    cboChequePosse.ItemData(cboChequePosse.NewIndex) = 1
    cboChequePosse.AddItem "2 Escritorio"
    cboChequePosse.ItemData(cboChequePosse.NewIndex) = 2
    cboChequePosse.AddItem "3 Pista"
    cboChequePosse.ItemData(cboChequePosse.NewIndex) = 3
End Sub
Private Sub cmd_data_b_Click()
    g_string = msk_data_baixa.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    msk_data_baixa.Text = RetiraGString(1)
    cmd_ok.SetFocus
    g_string = " "
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    msk_data.Text = RetiraGString(1)
    msk_valor.SetFocus
    g_string = " "
End Sub
Private Sub cmd_data_v_Click()
    g_string = msk_data_vencimento.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    msk_data_vencimento.Text = RetiraGString(1)
    optSituacao(0).SetFocus
    g_string = " "
End Sub
Private Sub cmd_excluir_Click()
    If Val(txt_numero.Text) > 0 Then
        If (MsgBox("Deseja Realmente Excluir Este Registro?", 4 + 32 + 256, "Exclusão de Registro!")) = 6 Then
            If ChequeFolha.Excluir(g_empresa, CDate(msk_data.Text), txt_numero.Text) Then
                LimpaTela
                If ChequeFolha.LocalizarUltimo(g_empresa) Then
                    AtualTela
                Else
                    DesativaBotoes
                    cmd_novo.Enabled = True
                    cmd_sair.Enabled = True
                    cmd_novo.SetFocus
                End If
            Else
                MsgBox "Não foi possível excluir este registro!", vbCritical, "Erro de Integridade!"
            End If
        End If
    End If
End Sub
Private Sub cmd_fornecedores_Click()
    cmd_fornecedores.FontBold = True
    cmd_funcionario.FontBold = False
'    dta_funcionario.RecordSource = "SELECT Codigo, Nome FROM Fornecedor WHERE Empresa = " & g_empresa & " ORDER BY Nome"
'    dta_funcionario.Refresh
    Set adodcFuncionario.Recordset = Conectar.RsConexao("SELECT Codigo, Nome FROM Fornecedor WHERE Empresa = " & g_empresa & " ORDER BY Nome")
    dtcboFuncionario.SetFocus
End Sub
Private Sub cmd_funcionario_Click()
    cmd_funcionario.FontBold = True
    cmd_fornecedores.FontBold = False
'    dta_funcionario.RecordSource = "SELECT Codigo, Nome FROM Funcionario WHERE Empresa = " & g_empresa & " AND Situacao = 'A' ORDER BY Nome"
'    dta_funcionario.Refresh
    Set adodcFuncionario.Recordset = Conectar.RsConexao("SELECT Codigo, Nome FROM Funcionario WHERE Empresa = " & g_empresa & " AND Situacao = " & preparaTexto("A") & " ORDER BY Nome")
    dtcboFuncionario.SetFocus
End Sub
Private Sub cmd_imprimir_Click()
    lCruzarCheque = False
    If MsgBox("Deseja Cruzar este cheque?", vbQuestion + vbYesNo + vbDefaultButton2, "Cruzar Cheque") = vbYes Then
        lCruzarCheque = True
    End If
    If Me.optChequeFolha Then
        If ConfiguracaoCheque.LocalizarCodigo(g_empresa, 1) Then
            If SelecionaImpressoraHP(Me) Then
                Call GravaAuditoria(1, Me.name, 7, "")
                RelatorioChequeFolha (False)
                cmd_novo.SetFocus
            End If
        Else
            MsgBox "Não existe configuração para impressão de cheque.", vbInformation, "Ajustar Configuração de Cheque!"
        End If
    Else
        If ConfiguracaoCheque.LocalizarCodigo(g_empresa, 1) Then
            If SelecionaImpressoraEpson(Me) Then
                Call GravaAuditoria(1, Me.name, 7, "")
                RelatorioChequeFormularioContinuo (False)
                cmd_novo.SetFocus
            End If
        Else
            MsgBox "Não existe configuração para impressão de cheque.", vbInformation, "Ajustar Configuração de Cheque!"
        End If
    End If
End Sub
Private Sub cmd_novo_Click()
    LimpaTela
    Inclui
    frm_dados.Enabled = True
    dtcboConta.SetFocus
End Sub
Private Sub cmd_ok_Click()
    On Error GoTo FileError
    If ValidaCampos Then
        AtivaBotoes
        If lOpcao = 1 Then
            AtualTabe
            If Not ChequeFolha.Incluir Then
                MsgBox "Não foi possível incluir este registro!", vbCritical, "Erro de Integridade!"
            Else
                l_numero = txt_numero.Text
                l_data = CDate(msk_data.Text)
            End If
        ElseIf lOpcao = 2 Then
            AtualTabe
            If Not ChequeFolha.Alterar(g_empresa, l_data, l_numero) Then
                MsgBox "Não foi possível alterar este registro!", vbCritical, "Erro de Integridade!"
            Else
                l_numero = txt_numero.Text
                l_data = CDate(msk_data.Text)
            End If
        End If
        If ChequeFolha.LocalizarCodigo(g_empresa, l_data, l_numero) Then
            AtualTela
            cmd_imprimir.SetFocus
        Else
            LimpaTela
        End If
    End If
    Exit Sub

FileError:
    MsgBox "Erro desconhecido ao atualizar banco de dados.", vbCritical, "Erro Não Identificado!"
    Exit Sub
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If dtcboConta.BoundText = "" Then
        MsgBox "Selecione uma conta.", vbInformation, "Atenção!"
        dtcboConta.SetFocus
    ElseIf Not txt_numero.Text <> "" Then
        MsgBox "Informe o número do cheque.", vbInformation, "Atenção!"
        txt_numero.SetFocus
    ElseIf Not IsDate(msk_data.Text) Then
        MsgBox "Informe a data de emissão.", vbInformation, "Atenção!"
        msk_data.SetFocus
    ElseIf Not fValidaValor2(msk_valor) > 0 And Not optSituacao(2) Then
        MsgBox "Informe o valor do cheque.", vbInformation, "Atenção!"
        msk_valor.SetFocus
    'ElseIf Not txt_nome.Text <> "" Then
    '    MsgBox "Informe o favorecido.", vbInformation, "Atenção!"
    '    txt_nome.SetFocus
    ElseIf Not txt_historico.Text <> "" Then
        MsgBox "Informe o historico.", vbInformation, "Atenção!"
        txt_historico.SetFocus
    ElseIf Not optSituacao(0) And Not optSituacao(1) And Not optSituacao(2) Then
        MsgBox "Escolha uma situação.", vbInformation, "Atenção!"
        optSituacao(0).SetFocus
    'ElseIf Not IsDate(msk_data_vencimento.Text) Then
    '    MsgBox "Informe a data de vencimento.", vbInformation, "Atenção!"
    '    msk_data_vencimento.SetFocus
    ElseIf CDate(msk_data_vencimento.Text) < CDate(msk_data.Text) Then
        MsgBox "Data de vencimento deve ser igual maior que " & msk_data & ".", vbInformation, "Atenção!"
        msk_data_vencimento.SetFocus
    ElseIf (optSituacao(1) Or optSituacao(2)) And Not IsDate(msk_data_baixa) Then
        MsgBox "Informe a data da baixa.", vbInformation, "Atenção!"
        msk_data_baixa.SetFocus
    ElseIf (optSituacao(1) Or optSituacao(2)) Then
        If CDate(msk_data_baixa) < CDate(msk_data) Then
            MsgBox "Data da baixa deve ser igual maior que " & msk_data & ".", vbInformation, "Atenção!"
            msk_data_baixa.SetFocus
        Else
            ValidaCampos = True
        End If
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_pesquisa_Click()
    consulta_cheque_folha.Show 1
    If Len(g_string) > 0 Then
        l_data = RetiraGString(1)
        l_numero = RetiraGString(2)
        If ChequeFolha.LocalizarCodigo(g_empresa, l_data, l_numero) Then
            AtualTela
        End If
    End If
End Sub
Private Sub cmd_primeiro_Click()
    If ChequeFolha.LocalizarPrimeiro() Then
        AtualTela
        cmd_proximo.SetFocus
    Else
        MsgBox "Não há registros nesta empresa.", vbInformation, "Atenção!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    If ChequeFolha.LocalizarProximo() Then
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
    If ChequeFolha.LocalizarUltimo(g_empresa) Then
        AtualTela
        cmd_anterior.SetFocus
    Else
        MsgBox "Não há registros nesta empresa.", vbInformation, "Atenção!"
    End If
End Sub
Private Sub dtcboConta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        If lOpcao > 0 Then
            msk_data.SetFocus
        Else
            'If MovBancario.LocalizarUltimo(g_empresa, dtcboConta.BoundText) Then
            '    AtualTela
            '    AtivaBotoes
            '    AtualTela
            '    cmd_anterior.SetFocus
            'Else
            '    DesativaBotoes
            '    cmd_novo.Enabled = True
            '    cmd_sair.Enabled = True
            '    LimpaTela
            '    MsgBox "Não há registros nesta conta.", vbInformation, "Erro de Verificação!"
            'End If
        End If
    End If
End Sub
Private Sub dtcboConta_LostFocus()
    If dtcboConta.BoundText <> "" Then
        If lOpcao = 1 Then
            txt_numero.Text = Format(ChequeFolha.ProximoNumeroCheque(g_empresa, dtcboConta.BoundText), "000000")
            If Not IsDate(msk_data.Text) Then
                msk_data.Text = Format(g_data_def, "dd/mm/yyyy")
                msk_data_vencimento.Text = msk_data.Text
            End If
            msk_valor.SetFocus
        End If
    End If
End Sub
Private Sub dtcboFuncionario_GotFocus()
    txt_nome.Visible = False
    dtcboFuncionario.BoundText = ""
End Sub
Private Sub dtcboFuncionario_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        txt_historico.SetFocus
    End If
End Sub
Private Sub dtcboFuncionario_LostFocus()
    txt_nome.Visible = True
    If dtcboFuncionario.BoundText <> "" Then
        txt_nome.Text = dtcboFuncionario.Text
        txt_historico.SetFocus
    Else
        txt_nome.Text = ""
        txt_nome.SetFocus
    End If
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    If flag_emissao_cheques_folhas = 0 Then
        DesativaBotoes
        If ChequeFolha.LocalizarUltimo(g_empresa) Then
            AtualTela
            AtivaBotoes
        Else
            cmd_novo.Enabled = True
            cmd_sair.Enabled = True
        End If
        cmd_novo.SetFocus
    Else
        flag_emissao_cheques_folhas = 0
    End If
    If g_nome_usuario Like "*rick*" Then
        cmd_funcionario.FontBold = True
        cmd_fornecedores.FontBold = False
'        dta_funcionario.RecordSource = "Select * From Funcionario Where Empresa = " & g_empresa & " And Situacao = 'A' Order By Nome"
'        dta_funcionario.Refresh
        Set adodcFuncionario.Recordset = Conectar.RsConexao("SELECT Codigo, Nome FROM Funcionario WHERE Empresa = " & g_empresa & " AND Situacao = " & preparaTexto("A") & " ORDER BY Nome")
    Else
        cmd_fornecedores.FontBold = True
        cmd_funcionario.FontBold = False
'        dta_funcionario.RecordSource = "Select * From Fornecedor Where Empresa = " & g_empresa & " Order By Nome"
'        dta_funcionario.Refresh
        Set adodcFuncionario.Recordset = Conectar.RsConexao("SELECT Codigo, Nome FROM Fornecedor WHERE Empresa = " & g_empresa & " ORDER BY Nome")
    End If
End Sub
Private Sub Form_Deactivate()
    flag_emissao_cheques_folhas = 1
End Sub
Private Sub Form_Load()
    Screen.MousePointer = 1
    CentraForm Me
    
'    adodc_conta.ConnectionString = gConnectionString
'    adodc_conta.RecordSource = "SELECT Codigo, Nome FROM Conta_Bancaria WHERE Empresa = " & g_empresa & " ORDER BY Nome"
'    adodc_conta.Refresh
    Set adodcConta.Recordset = Conectar.RsConexao("SELECT Codigo, Nome FROM PortadorFinanceiro WHERE (Empresa = " & g_empresa & " OR Empresa = 0) AND [Instituicao Financeira] = " & preparaBooleano(True) & " ORDER BY Nome")
    PreencheCboChequePosse

    'SelecionaImpressora
End Sub
Private Sub SelecionaImpressora()
Dim Impressora As Printer
    For Each Impressora In Printers
        If Impressora.DeviceName = "HP DeskJet 600" Then
            Set Printer = Impressora
            Exit Sub
        ElseIf Impressora.DeviceName = "HP 820CXI em PC_486-DX4" Then
            Set Printer = Impressora
            Exit Sub
        End If
    Next
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_baixa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub msk_data_GotFocus()
    If Not IsDate(msk_data.Text) Then
        msk_data.Text = Format(g_data_def, "dd/mm/yyyy")
    End If
End Sub
Private Sub msk_data_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_valor.SetFocus
    End If
End Sub
Private Sub msk_data_LostFocus()
    If IsDate(msk_data.Text) Then
        g_data_def = msk_data.Text
    End If
    msk_data_vencimento.Text = msk_data.Text
End Sub
Private Sub msk_data_vencimento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub msk_valor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    End If
    If KeyAscii = 13 Then
        KeyAscii = 0
        dtcboFuncionario.SetFocus
    End If
End Sub
Private Sub msk_valor_LostFocus()
    msk_valor.Text = Format(msk_valor.Text, "###,##0.00")
    If fValidaValor2(msk_valor.Text) > 0 Then
        lbl_extenso.Caption = FazExtenso(fValidaValor2(msk_valor.Text))
    Else
        lbl_extenso.Caption = ""
    End If
End Sub
Private Sub optSituacao_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data_baixa.SetFocus
    End If
End Sub
Private Sub optSituacao_LostFocus(Index As Integer)
    If optSituacao(0) Then
        msk_data_baixa.Text = "__/__/____"
    End If
End Sub
Private Sub optChequeFolha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_valor.SetFocus
    End If
End Sub
Private Sub optFormularioContinuo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_valor.SetFocus
    End If
End Sub
Private Sub txt_historico_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data_vencimento.SetFocus
        If lOpcao = 1 Then
            cmd_ok.SetFocus
        End If
    End If
End Sub
Private Sub txt_nome_Click()
    dtcboFuncionario.SetFocus
End Sub
Private Sub txt_nome_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_historico.SetFocus
    End If
End Sub
Private Sub txt_numero_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
