VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form cadastro_produto 
   Caption         =   "Cadastro de Produtos"
   ClientHeight    =   6525
   ClientLeft      =   120
   ClientTop       =   1935
   ClientWidth     =   6915
   Icon            =   "cad_produto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "cad_produto.frx":030A
   ScaleHeight     =   6525
   ScaleWidth      =   6915
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   60
      Picture         =   "cad_produto.frx":0750
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "Cria um novo registro."
      Top             =   5580
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   960
      Picture         =   "cad_produto.frx":1DE2
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Altera o registro atual."
      Top             =   5580
      Width           =   795
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   1860
      Picture         =   "cad_produto.frx":32DC
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Exclui o registro atual."
      Top             =   5580
      Width           =   795
   End
   Begin VB.CommandButton cmd_pesquisa 
      Caption         =   "&Pesquisa"
      Height          =   855
      Left            =   2760
      Picture         =   "cad_produto.frx":496E
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "Pesquisa um registro específico."
      Top             =   5580
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3660
      Picture         =   "cad_produto.frx":5DE0
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   5580
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Enabled         =   0   'False
      Height          =   5415
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6795
      Begin VB.CheckBox chkExclusivoPosto 
         Caption         =   "Exclusivo Posto"
         Height          =   255
         Left            =   4680
         TabIndex        =   21
         Top             =   2760
         Width           =   1575
      End
      Begin VB.CheckBox chkExclusivoLoja 
         Caption         =   "Exclusivo Loja"
         Height          =   255
         Left            =   4680
         TabIndex        =   28
         Top             =   3480
         Width           =   1575
      End
      Begin VB.TextBox txt_codigo_barra 
         Height          =   285
         Left            =   5040
         MaxLength       =   13
         TabIndex        =   4
         Top             =   240
         Width           =   1635
      End
      Begin VB.TextBox txt_margem_lucro_desejada 
         Height          =   285
         Left            =   5580
         MaxLength       =   6
         TabIndex        =   18
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CheckBox chk_inativo 
         Caption         =   "Produto &Inativo"
         Height          =   255
         Left            =   4680
         TabIndex        =   31
         Top             =   3900
         Width           =   1575
      End
      Begin VB.CheckBox chk_comissao 
         Caption         =   "&Produto c/ comissão"
         Height          =   255
         Left            =   4680
         TabIndex        =   34
         Top             =   4320
         Width           =   1995
      End
      Begin VB.TextBox txt_quantidade_maxima 
         Height          =   285
         Left            =   5580
         MaxLength       =   13
         TabIndex        =   25
         Top             =   3120
         Width           =   1095
      End
      Begin VB.TextBox txt_preco_venda 
         Height          =   285
         Left            =   1740
         MaxLength       =   13
         TabIndex        =   20
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox txt_unidade 
         Height          =   285
         Left            =   1740
         MaxLength       =   3
         TabIndex        =   8
         Top             =   960
         Width           =   435
      End
      Begin VB.TextBox txt_codigo 
         Height          =   285
         Left            =   1740
         MaxLength       =   4
         TabIndex        =   2
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txt_nome 
         Height          =   285
         Left            =   1740
         MaxLength       =   40
         TabIndex        =   6
         Top             =   600
         Width           =   4935
      End
      Begin VB.TextBox txt_quantidade_minima 
         Height          =   285
         Left            =   1740
         MaxLength       =   13
         TabIndex        =   23
         Top             =   3120
         Width           =   1095
      End
      Begin VB.TextBox txt_preco_custo 
         Height          =   285
         Left            =   1740
         MaxLength       =   13
         TabIndex        =   16
         Top             =   2400
         Width           =   1095
      End
      Begin MSMask.MaskEdBox msk_data_ultima_compra 
         Height          =   300
         Left            =   1740
         TabIndex        =   30
         Top             =   3840
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_data_acerto 
         Height          =   300
         Left            =   1740
         TabIndex        =   33
         Top             =   4260
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSAdodcLib.Adodc adodcGrupo 
         Height          =   330
         Left            =   2760
         Top             =   1320
         Visible         =   0   'False
         Width           =   3375
         _ExtentX        =   5953
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
         Caption         =   "adodcGrupo"
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
      Begin MSDataListLib.DataCombo dtcboGrupo 
         Bindings        =   "cad_produto.frx":7472
         Height          =   315
         Left            =   1740
         TabIndex        =   10
         Top             =   1320
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   "dtcboGrupo"
      End
      Begin MSAdodcLib.Adodc adodcSubGrupo 
         Height          =   330
         Left            =   2760
         Top             =   1680
         Visible         =   0   'False
         Width           =   3375
         _ExtentX        =   5953
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
         Caption         =   "adodcSubGrupo"
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
      Begin MSDataListLib.DataCombo dtcboSubGrupo 
         Bindings        =   "cad_produto.frx":748B
         Height          =   315
         Left            =   1740
         TabIndex        =   12
         Top             =   1680
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   "dtcboSubGrupo"
      End
      Begin MSAdodcLib.Adodc adodcAliquota 
         Height          =   330
         Left            =   2760
         Top             =   2040
         Visible         =   0   'False
         Width           =   3375
         _ExtentX        =   5953
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
         Caption         =   "adodcAliquota"
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
      Begin MSDataListLib.DataCombo dtcboAliquota 
         Bindings        =   "cad_produto.frx":74A7
         Height          =   315
         Left            =   1740
         TabIndex        =   14
         Top             =   2040
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   "dtcboAliquota"
      End
      Begin MSAdodcLib.Adodc adodcTCS 
         Height          =   330
         Left            =   2760
         Top             =   4620
         Visible         =   0   'False
         Width           =   3375
         _ExtentX        =   5953
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
         Caption         =   "adodcTCS"
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
      Begin MSDataListLib.DataCombo dtcboTCS 
         Bindings        =   "cad_produto.frx":74C3
         Height          =   315
         Left            =   1740
         TabIndex        =   36
         Top             =   4620
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   "dtcboTCS"
      End
      Begin MSAdodcLib.Adodc adodc_combustivel 
         Height          =   330
         Left            =   2760
         Top             =   4980
         Visible         =   0   'False
         Width           =   3375
         _ExtentX        =   5953
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
         Caption         =   "adodc_combustivel"
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
      Begin MSDataListLib.DataCombo dtcbo_combustivel 
         Bindings        =   "cad_produto.frx":74DA
         Height          =   315
         Left            =   1740
         TabIndex        =   38
         Top             =   4980
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   "dtcbo_combustivel"
      End
      Begin VB.Label Label10 
         Caption         =   "Co&mbustível"
         Height          =   300
         Left            =   120
         TabIndex        =   37
         Top             =   4980
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Código TC Smart"
         Height          =   315
         Index           =   7
         Left            =   120
         TabIndex        =   35
         Top             =   4620
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "Código de Barra"
         Height          =   315
         Left            =   3660
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Margem de Lucro Desejada"
         Height          =   315
         Left            =   3480
         TabIndex        =   17
         Top             =   2400
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "Sub-Grupo"
         Height          =   315
         Index           =   6
         Left            =   120
         TabIndex        =   11
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Alíquo&ta"
         Height          =   315
         Index           =   5
         Left            =   120
         TabIndex        =   13
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label lbl_quantidade_estoque 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1740
         TabIndex        =   27
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Quantidade Estoque"
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   26
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Quantidade &Máxima"
         Height          =   315
         Index           =   3
         Left            =   3960
         TabIndex        =   24
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Preço de &Venda"
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "&Unidade"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Código do P&roduto"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "&Grupo"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "&Nome"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "&Data última Compra"
         Height          =   315
         Left            =   120
         TabIndex        =   29
         Top             =   3840
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Data &Acerto"
         Height          =   315
         Left            =   120
         TabIndex        =   32
         Top             =   4260
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "&Quantidade Mínima"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Preço de C&usto"
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   2400
         Width           =   1575
      End
   End
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   4680
      TabIndex        =   46
      Top             =   5460
      Width           =   2175
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "cad_produto.frx":74FA
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "cad_produto.frx":8A7C
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "cad_produto.frx":9EEE
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "cad_produto.frx":B3E8
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   6060
      Picture         =   "cad_produto.frx":C8E2
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "Cancela o registro atual."
      Top             =   5580
      Width           =   795
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   5160
      Picture         =   "cad_produto.frx":DDDC
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Confirma o registro atual."
      Top             =   5580
      Width           =   795
   End
End
Attribute VB_Name = "cadastro_produto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lFlagCadastro As Integer
Dim lOpcao As Integer
Dim lCodigo As Long
Dim lEmpresa As Integer

Private Bomba As New cBomba
Private Estoque As New cEstoque
Private Produto As New cProduto
Private SubEstoque As New cSubEstoque
Private TicketCarDePara As New cTicketCarDePara
Private Sub AtivaBotoes()
    cmd_novo.Enabled = True
    cmd_alterar.Enabled = True
    cmd_pesquisa.Enabled = True
    cmd_excluir.Enabled = True
    cmd_sair.Enabled = True
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
    frm_move.Visible = True
    txt_codigo.Enabled = True
End Sub
Private Sub AtualTabe()
    Produto.CodigoGrupo = Val(dtcboGrupo.BoundText)
    Produto.Codigo = CLng(txt_codigo.Text)
    Produto.CodigoBarra = txt_codigo_barra.Text
    Produto.Nome = txt_nome.Text
    Produto.Unidade = txt_unidade.Text
    Produto.PrecoCusto = fValidaValor(txt_preco_custo.Text)
    Produto.MargemLucroDesejada = fValidaValor(txt_margem_lucro_desejada.Text)
    Produto.PrecoVenda = fValidaValor(txt_preco_venda.Text)
    Produto.EstoqueMinimo = fValidaValor(txt_quantidade_minima.Text)
    Produto.EstoqueMaximo = fValidaValor(txt_quantidade_maxima.Text)
    If IsDate(msk_data_ultima_compra.Text) Then
        Produto.DataUltimaCompra = msk_data_ultima_compra.Text
    Else
        Produto.DataUltimaCompra = "00:00:00"
    End If
    If IsDate(msk_data_acerto.Text) Then
        Produto.DataAcerto = msk_data_acerto.Text
    Else
        Produto.DataAcerto = "00:00:00"
    End If
    If lOpcao = 1 Then
        Produto.PrecoCustoAnterior = 0
        Produto.PrecoCustoMedio = 0
        Produto.EstoqueFisico = 0
    End If
    Produto.Inativo = chk_inativo.Value
    Produto.Comissao = chk_comissao.Value
    Produto.CodigoAliquota = Val(dtcboAliquota.BoundText)
    Produto.CodigoSubGrupo = Val(dtcboSubGrupo.BoundText)
    Produto.ExclusivoPosto = chkExclusivoPosto.Value
    Produto.ExclusivoLoja = chkExclusivoLoja.Value
    Produto.TipoCombustivel = dtcbo_combustivel.BoundText
    
    Bomba.PrecoCusto = fValidaValor(txt_preco_custo.Text)
    Bomba.PrecoVenda = fValidaValor(txt_preco_venda.Text)
End Sub
Private Sub AtualTela()
    lCodigo = Produto.Codigo
    dtcboGrupo.BoundText = Val(Produto.CodigoGrupo)
    dtcboSubGrupo.BoundText = Val(Produto.CodigoSubGrupo)
    dtcboAliquota.BoundText = Val(Produto.CodigoAliquota)
    txt_codigo.Text = Format(Produto.Codigo, "#000")
    txt_codigo_barra.Text = Produto.CodigoBarra
    txt_nome.Text = Produto.Nome
    txt_unidade.Text = Produto.Unidade
    txt_preco_custo.Text = Format(Produto.PrecoCusto, "####,##0.0000")
    txt_margem_lucro_desejada.Text = Format(Produto.MargemLucroDesejada, "##0.00")
    txt_preco_venda.Text = Format(Produto.PrecoVenda, "####,##0.0000")
    txt_quantidade_minima.Text = Format(Produto.EstoqueMinimo, "##,###,##0.00")
    txt_quantidade_maxima.Text = Format(Produto.EstoqueMaximo, "##,###,##0.00")
    If Estoque.LocalizarCodigo(g_empresa, Produto.Codigo) Then
        lbl_quantidade_estoque = Format(Estoque.Quantidade, "##,###,##0.00")
        If Estoque.PrecoVenda > 0 Then
            txt_preco_venda = Format(Estoque.PrecoVenda, "####,##0.0000")
        End If
    Else
        MsgBox "Quantidade não cadastrada", vbInformation, "Erro de Integridade!"
        lbl_quantidade_estoque = "0"
    End If
    If Produto.DataUltimaCompra <> "00:00:00" Then
        msk_data_ultima_compra.Text = Format(Produto.DataUltimaCompra, "dd/mm/yyyy")
    Else
        msk_data_ultima_compra.Text = "__/__/____"
    End If
    If Produto.DataAcerto <> "00:00:00" Then
        msk_data_acerto.Text = Format(Produto.DataAcerto, "dd/mm/yyyy")
    Else
        msk_data_acerto.Text = "__/__/____"
    End If
    If Produto.Comissao Then
        chk_comissao.Value = 1
    Else
        chk_comissao.Value = 0
    End If
    If Produto.Inativo Then
        chk_inativo.Value = 1
    Else
        chk_inativo.Value = 0
    End If
    If Produto.ExclusivoPosto Then
        chkExclusivoPosto.Value = 1
    Else
        chkExclusivoPosto.Value = 0
    End If
    If Produto.ExclusivoLoja Then
        chkExclusivoLoja.Value = 1
    Else
        chkExclusivoLoja.Value = 0
    End If
    dtcboTCS.BoundText = ""
    If TicketCarDePara.LocalizarCodigo(Produto.Codigo) Then
        dtcboTCS.BoundText = Val(TicketCarDePara.CodigoTCS)
    End If
    dtcbo_combustivel.BoundText = Produto.TipoCombustivel
    If dtcboGrupo.BoundText = 4 Then
        dtcbo_combustivel.Enabled = True
    Else
        dtcbo_combustivel.Enabled = False
    End If
    
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
    Set Bomba = Nothing
    Set Estoque = Nothing
    Set Produto = Nothing
    Set SubEstoque = Nothing
    Set TicketCarDePara = Nothing
    frm_cadastro.Show
End Sub
Private Sub Inclui()
    lOpcao = 1
    DesativaBotoes
    cmd_novo.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    txt_codigo.Text = 0
    If Produto.LocalizarUltimo Then
        txt_codigo.Text = Produto.Codigo
    End If
    txt_codigo.Text = Format(CLng(txt_codigo.Text) + 1, "#000")
End Sub
Private Sub chk_comissao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dtcboTCS.SetFocus
    End If
End Sub
Private Sub chk_inativo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data_acerto.SetFocus
    End If
End Sub
Private Sub chkExclusivoLoja_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data_ultima_compra.SetFocus
    End If
End Sub
Private Sub chkExclusivoPosto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_quantidade_minima.SetFocus
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
    txt_codigo.Enabled = False
    txt_nome.SetFocus
End Sub
Private Sub cmd_anterior_Click()
    If Produto.LocalizarAnterior Then
        AtualTela
    Else
        MsgBox "Início de Arquivo.", vbInformation, "Atenção!"
        cmd_proximo.SetFocus
    End If
End Sub
Private Sub cmd_cancelar_Click()
    LimpaTela
    lOpcao = 0
    If Produto.LocalizarCodigo(lCodigo) Then
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
    dtcboGrupo.BoundText = ""
    dtcboSubGrupo.BoundText = ""
    dtcboAliquota.BoundText = ""
    txt_codigo.Text = ""
    txt_codigo_barra.Text = ""
    txt_nome.Text = ""
    txt_unidade.Text = ""
    txt_preco_custo.Text = ""
    txt_margem_lucro_desejada.Text = ""
    txt_preco_venda.Text = ""
    txt_quantidade_minima.Text = ""
    txt_quantidade_maxima.Text = ""
    lbl_quantidade_estoque.Caption = Format(0, "###,##0.00")
    msk_data_ultima_compra.Text = "__/__/____"
    msk_data_acerto.Text = "__/__/____"
    chk_comissao.Value = 0
    chkExclusivoPosto.Value = 1
    chkExclusivoLoja.Value = 0
    If UCase(g_nome_empresa) Like "*LOJA*" Then
        chkExclusivoPosto.Value = 0
        chkExclusivoLoja.Value = 1
    End If
    dtcboTCS.BoundText = ""
    dtcbo_combustivel.BoundText = ""
    dtcbo_combustivel.Enabled = False
End Sub
Function ProdutoInativo() As Boolean
    Dim i As Integer
    ProdutoInativo = True
    If chk_inativo = 0 Then
        Exit Function
    End If
    i = 0
    Do Until i = 9
        i = i + 1
        If i = 7 Then
            i = 9
        End If
        If Estoque.LocalizarCodigo(i, Val(txt_codigo.Text)) Then
            If Estoque.Quantidade > 0 Then
                MsgBox "Este produto não pode ser inativado." & Chr(10) & Chr(13) & Chr(10) & "Este produto tem " & Estoque.Quantidade & " " & txt_unidade.Text & " na empresa: " & i & ".", vbInformation, "Mensagem do Sistema"
                ProdutoInativo = False
                Exit Do
            End If
        End If
    Loop
End Function
Private Sub cmd_excluir_Click()
    If Val(txt_codigo.Text) > 0 Then
        If (MsgBox("Deseja excluir este registro?", 4 + 32 + 256, "Exclusão de Registro!")) = 6 Then
            If Produto.Excluir(Val(txt_codigo.Text)) Then
                LimpaTela
                If Produto.LocalizarUltimo Then
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
'
'    Exit Sub
'    zzCriaEstoque
'    Exit Sub
'
    LimpaTela
    Inclui
    frm_dados.Enabled = True
    txt_codigo_barra.SetFocus
End Sub
Private Sub cmd_ok_Click()
Dim xEmpresa As Integer
    On Error GoTo FileError
    If ValidaCampos Then
        AtivaBotoes
        If lOpcao = 1 Then
            AtualTabe
            If Produto.Incluir Then
                lCodigo = Val(txt_codigo.Text)
            Else
                MsgBox "Não foi possível incluir este registro!", vbInformation, "Erro de Verificação!"
            End If
        ElseIf lOpcao = 2 Then
            xEmpresa = g_empresa
            If (MsgBox("Deseja regravar o preço de venda em todas empresas?", vbYesNo + vbDefaultButton2 + vbQuestion, "Alteração de Preços")) = vbYes Then
                xEmpresa = 0
            End If
            AtualTabe
            If Not Produto.Alterar(lCodigo, xEmpresa) Then
                MsgBox "Não foi possível alterar este registro!", vbInformation, "Erro de Verificação!"
            End If
            If ReadINI("ESTOQUE", "Atualizar Preco no Cadastro de Bomba", ArqSgpIni) = "SIM" Then
                Call Bomba.AlterarCodigoProduto(g_empresa, Val(txt_codigo.Text))
            End If
        End If
        lOpcao = 0
        Call Produto.LocalizarCodigo(lCodigo)
        cmd_novo.SetFocus
    End If
    Exit Sub
FileError:
    MsgBox Error
    Exit Sub
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If txt_codigo.Text = "" Then
        MsgBox "Informe o código do produto.", vbInformation, "Atenção!"
        txt_codigo.SetFocus
    ElseIf txt_nome.Text = "" Then
        MsgBox "Informe o nome do produto.", vbInformation, "Atenção!"
        txt_nome.SetFocus
    ElseIf txt_unidade.Text = "" Then
        MsgBox "Informe a unidade.", vbInformation, "Atenção!"
        txt_unidade.SetFocus
    ElseIf Val(dtcboGrupo.BoundText) = 0 Then
        MsgBox "Escolha um grupo.", vbInformation, "Atenção!"
        dtcboGrupo.SetFocus
    ElseIf Val(dtcboSubGrupo.BoundText) = 0 Then
        MsgBox "Selecione um Sub-Grupo.", vbInformation, "Atenção!"
        dtcboSubGrupo.SetFocus
    ElseIf Not fValidaValor4(txt_preco_custo.Text) > 0 Then
        MsgBox "Informe o preço de custo.", vbInformation, "Atenção!"
        txt_preco_custo.SetFocus
    ElseIf Not fValidaValor(txt_margem_lucro_desejada.Text) > 0 Then
        MsgBox "Informe a margem de lucro desejada.", vbInformation, "Atenção!"
        txt_margem_lucro_desejada.SetFocus
    ElseIf Not fValidaValor4(txt_preco_venda.Text) > 0 Then
        MsgBox "Informe o preço de venda.", vbInformation, "Atenção!"
        txt_preco_venda.SetFocus
    'ElseIf Not fValidaValor2(txt_quantidade_minima.Text) > 0 Then
    '    MsgBox "Informe a quantidade mínima.", vbInformation, "Atenção!"
    '    txt_quantidade_minima.SetFocus
    'ElseIf Not fValidaValor2(txt_quantidade_maxima.Text) > 0 Then
    '    MsgBox "Informe a quantidade máxima.", vbInformation, "Atenção!"
    '    txt_quantidade_maxima.SetFocus
    ElseIf Val(dtcboAliquota.BoundText) = 0 Then
        MsgBox "Informe a aliquota.", vbInformation, "Atenção!"
        dtcboAliquota.SetFocus
    ElseIf Not ProdutoInativo Then
        chk_inativo.SetFocus
'    ElseIf Not IsDate(msk_data_acerto) Then
'        MsgBox "Informe a data do acerto.", vbInformation, "Atenção!"
'        msk_data_acerto.SetFocus
    ElseIf Val(dtcboGrupo.BoundText) = 4 And dtcbo_combustivel.BoundText = "" Then
        MsgBox "Selecione o tipo de combustível.", vbInformation, "Atenção!"
        dtcbo_combustivel.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_pesquisa_Click()
    consulta_produto.Show 1
    If Len(g_string) > 0 Then
        lCodigo = RetiraGString(1)
        Call Produto.LocalizarCodigo(lCodigo)
        AtualTela
    End If
End Sub
Private Sub cmd_primeiro_Click()
    If Produto.LocalizarPrimeiro Then
        AtualTela
        cmd_proximo.SetFocus
    Else
        MsgBox "Não há registro.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    If Produto.LocalizarProximo Then
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
    If Produto.LocalizarUltimo Then
        AtualTela
        cmd_anterior.SetFocus
    Else
        MsgBox "Não há registro.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub dtcbo_combustivel_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub dtcboaliquota_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        txt_preco_custo.SetFocus
    End If
End Sub
Private Sub dtcbogrupo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        dtcboSubGrupo.SetFocus
    End If
End Sub
Private Sub dtcboGrupo_LostFocus()
    If dtcboGrupo.BoundText = 4 Then
        dtcbo_combustivel.Enabled = True
    Else
        dtcbo_combustivel.BoundText = ""
        dtcbo_combustivel.Enabled = False
    End If
End Sub
Private Sub dtcbosubgrupo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        dtcboAliquota.SetFocus
    End If
End Sub
Private Sub dtcboTCS_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        If dtcbo_combustivel.Enabled Then
            dtcbo_combustivel.SetFocus
        Else
            cmd_ok.SetFocus
        End If
    End If
End Sub
Private Sub Form_Activate()
    If g_empresa <> lEmpresa Then
        lFlagCadastro = 0
    End If
    If lFlagCadastro = 0 Then
        lOpcao = 0
        lEmpresa = g_empresa
        DesativaBotoes
        If Produto.LocalizarUltimo Then
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
    
    adodcAliquota.ConnectionString = gConnectionString
    adodcAliquota.RecordSource = "SELECT Codigo, Nome FROM Aliquota ORDER BY Nome"
    adodcAliquota.Refresh
    adodcGrupo.ConnectionString = gConnectionString
    adodcGrupo.RecordSource = "SELECT Codigo, Nome FROM Grupo ORDER BY Nome"
    adodcGrupo.Refresh
    adodcSubGrupo.ConnectionString = gConnectionString
    adodcSubGrupo.RecordSource = "SELECT Codigo, Nome FROM Sub_Grupo ORDER BY Nome"
    adodcSubGrupo.Refresh
    adodcTCS.ConnectionString = gConnectionString
    adodcTCS.RecordSource = "SELECT Codigo, Nome FROM TicketCarProduto ORDER BY Nome"
    adodcTCS.Refresh
    adodc_combustivel.ConnectionString = gConnectionString
    adodc_combustivel.RecordSource = "SELECT Codigo, Nome FROM Combustivel WHERE Empresa = " & g_empresa & " ORDER BY Nome"
    adodc_combustivel.Refresh

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_acerto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        chk_comissao.SetFocus
    End If
End Sub
Private Sub msk_data_ultima_compra_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        chk_inativo.SetFocus
    End If
End Sub
Private Sub txt_codigo_barra_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_nome.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_nome.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_codigo_LostFocus()
    txt_codigo = Format(txt_codigo, "#000")
    If lOpcao = 1 And txt_codigo <> "" Then
        If Produto.LocalizarCodigo(Val(txt_codigo.Text)) Then
            MsgBox "Já existe produto cadastrado com este código." & Chr(10) & Chr(10) & "Mude o código informado.", vbInformation, "Duplicidade de Registro!"
            txt_codigo.SetFocus
        End If
    End If
End Sub
Private Sub txt_margem_lucro_desejada_GotFocus()
    txt_margem_lucro_desejada.SelStart = 0
    txt_margem_lucro_desejada.SelLength = Len(txt_margem_lucro_desejada.Text)
End Sub
Private Sub txt_margem_lucro_desejada_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_preco_venda.SetFocus
    End If
End Sub
Private Sub txt_margem_lucro_desejada_LostFocus()
Dim xValor As Currency
    txt_margem_lucro_desejada.Text = Format(txt_margem_lucro_desejada.Text, "##0.00")
    If fValidaValor(txt_margem_lucro_desejada.Text) > 0 Then
        xValor = Format(fValidaValor(txt_preco_custo.Text) + (fValidaValor(txt_preco_custo.Text) * fValidaValor(txt_margem_lucro_desejada.Text) / 100), "00000000.00")
        If fValidaValor(txt_preco_venda.Text) <> xValor Then
            If (MsgBox("De acordo com a margem de lucro desejada em " & txt_margem_lucro_desejada.Text & "%." & Chr(10) & "O preço de venda deveria ser " & Format(xValor, "###,##0.00") & "." & Chr(10) & "E o preço atual de venda é " & txt_preco_venda.Text & "." & Chr(10) & Chr(10) & "Deseja alterá-lo?", vbQuestion + vbYesNo + vbDefaultButton2, "Preço de Venda Fora da Margem Desejada!")) = vbYes Then
                txt_preco_venda.Text = Format(xValor, "###,##0.0000")
            End If
        End If
    End If
End Sub
Private Sub txt_nome_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_unidade.SetFocus
    End If
End Sub
Private Sub txt_nome_LostFocus()
    If lOpcao = 1 And txt_nome <> "" Then
        If Produto.LocalizarNome(txt_nome.Text) Then
            If (MsgBox("Já existe produto cadastrado com este nome." & Chr(10) & Chr(10) & "Deseja cadastrar assim mesmo?", 4 + 32 + 256, "Duplicidade de Registro!")) = 7 Then
                txt_nome.SetFocus
            End If
        End If
    End If
End Sub
Private Sub txt_preco_custo_GotFocus()
    txt_preco_custo.SelStart = 0
    txt_preco_custo.SelLength = Len(txt_preco_custo.Text)
End Sub
Private Sub txt_preco_custo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_margem_lucro_desejada.SetFocus
    End If
End Sub
Private Sub txt_preco_custo_LostFocus()
    txt_preco_custo.Text = Format(txt_preco_custo.Text, "####,##0.0000")
End Sub
Private Sub txt_preco_venda_GotFocus()
    txt_preco_venda.SelStart = 0
    txt_preco_venda.SelLength = Len(txt_preco_venda.Text)
End Sub
Private Sub txt_preco_venda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        chkExclusivoPosto.SetFocus
    End If
End Sub
Private Sub txt_preco_venda_LostFocus()
    txt_preco_venda.Text = Format(txt_preco_venda.Text, "####,##0.0000")
End Sub
Private Sub txt_quantidade_maxima_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        chkExclusivoLoja.SetFocus
    End If
End Sub
Private Sub txt_quantidade_maxima_LostFocus()
    txt_quantidade_maxima = Format(txt_quantidade_maxima, "##,###,##0.00")
End Sub
Private Sub txt_quantidade_minima_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_quantidade_maxima.SetFocus
    End If
End Sub
Private Sub txt_quantidade_minima_LostFocus()
    txt_quantidade_minima = Format(txt_quantidade_minima, "##,###,##0.00")
End Sub
Private Sub txt_unidade_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dtcboGrupo.SetFocus
    End If
End Sub
