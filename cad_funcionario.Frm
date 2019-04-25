VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form cadastro_funcionario 
   Caption         =   "Cadastro de Funcionários"
   ClientHeight    =   7215
   ClientLeft      =   1725
   ClientTop       =   1155
   ClientWidth     =   10230
   Icon            =   "cad_funcionario.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "cad_funcionario.frx":030A
   ScaleHeight     =   7215
   ScaleWidth      =   10230
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   60
      Picture         =   "cad_funcionario.frx":0750
      Style           =   1  'Graphical
      TabIndex        =   108
      ToolTipText     =   "Cria um novo registro."
      Top             =   6240
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   960
      Picture         =   "cad_funcionario.frx":1DE2
      Style           =   1  'Graphical
      TabIndex        =   109
      ToolTipText     =   "Altera o registro atual."
      Top             =   6240
      Width           =   795
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   1860
      Picture         =   "cad_funcionario.frx":32DC
      Style           =   1  'Graphical
      TabIndex        =   110
      ToolTipText     =   "Exclui o registro atual."
      Top             =   6240
      Width           =   795
   End
   Begin VB.CommandButton cmd_pesquisa 
      Caption         =   "&Pesquisa"
      Height          =   855
      Left            =   2760
      Picture         =   "cad_funcionario.frx":496E
      Style           =   1  'Graphical
      TabIndex        =   111
      ToolTipText     =   "Pesquisa um registro específico."
      Top             =   6240
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3660
      Picture         =   "cad_funcionario.frx":5DE0
      Style           =   1  'Graphical
      TabIndex        =   112
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   6240
      Width           =   795
   End
   Begin TabDlg.SSTab tab_dados 
      Height          =   6135
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   10821
      _Version        =   393216
      TabHeight       =   529
      TabCaption(0)   =   "Dados Pessoais"
      TabPicture(0)   =   "cad_funcionario.frx":7472
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label12"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frm_dados_1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Referência Pessoal"
      TabPicture(1)   =   "cad_funcionario.frx":748E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frm_dados_2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Dados para Folha de Pagamento"
      TabPicture(2)   =   "cad_funcionario.frx":74AA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "frm_dados_3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin Threed.SSFrame frm_dados_1 
         Height          =   5655
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   9855
         _Version        =   65536
         _ExtentX        =   17383
         _ExtentY        =   9975
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
         Begin VB.TextBox txt_cep 
            Height          =   285
            Left            =   1680
            MaxLength       =   8
            TabIndex        =   15
            Top             =   1980
            Width           =   1215
         End
         Begin VB.TextBox txt_uf 
            Height          =   285
            Left            =   7200
            MaxLength       =   2
            TabIndex        =   13
            Top             =   1620
            Width           =   375
         End
         Begin VB.TextBox txt_cidade 
            Height          =   285
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   11
            Top             =   1620
            Width           =   2535
         End
         Begin VB.TextBox txt_codigo 
            Height          =   285
            Left            =   1680
            MaxLength       =   3
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   180
            Width           =   615
         End
         Begin VB.TextBox txt_nome 
            Height          =   285
            Left            =   1680
            MaxLength       =   40
            TabIndex        =   5
            Top             =   540
            Width           =   4815
         End
         Begin VB.TextBox txt_endereco 
            Height          =   285
            Left            =   1680
            MaxLength       =   40
            TabIndex        =   7
            Top             =   900
            Width           =   4815
         End
         Begin VB.TextBox txt_bairro 
            Height          =   285
            Left            =   1680
            MaxLength       =   30
            TabIndex        =   9
            Top             =   1260
            Width           =   3735
         End
         Begin VB.TextBox txt_cpf 
            Height          =   285
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   23
            Top             =   2700
            Width           =   2535
         End
         Begin VB.TextBox txt_identidade 
            Height          =   285
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   19
            Top             =   2340
            Width           =   2535
         End
         Begin VB.TextBox txt_cargo 
            Height          =   285
            Left            =   7200
            MaxLength       =   20
            TabIndex        =   25
            Top             =   2700
            Width           =   2535
         End
         Begin VB.TextBox txt_situacao 
            Height          =   285
            Left            =   1680
            MaxLength       =   1
            TabIndex        =   27
            Top             =   3060
            Width           =   255
         End
         Begin VB.TextBox txt_pis 
            Height          =   285
            Left            =   7200
            MaxLength       =   20
            TabIndex        =   29
            Top             =   3060
            Width           =   2535
         End
         Begin VB.TextBox txt_carteira_trabalho 
            Height          =   285
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   31
            Top             =   3420
            Width           =   2535
         End
         Begin VB.TextBox txt_cnh 
            Height          =   285
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   35
            Top             =   3780
            Width           =   2535
         End
         Begin VB.TextBox txt_titulo_eleitor 
            Height          =   285
            Left            =   7200
            MaxLength       =   20
            TabIndex        =   37
            Top             =   3780
            Width           =   2535
         End
         Begin VB.TextBox txt_reservista 
            Height          =   285
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   39
            Top             =   4140
            Width           =   2535
         End
         Begin VB.TextBox txt_sexo 
            Height          =   285
            Left            =   7200
            MaxLength       =   1
            TabIndex        =   41
            Top             =   4140
            Width           =   255
         End
         Begin VB.TextBox txt_grau_instrucao 
            Height          =   285
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   43
            Top             =   4500
            Width           =   2535
         End
         Begin VB.TextBox txt_estado_civil 
            Height          =   285
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   51
            Top             =   5220
            Width           =   2535
         End
         Begin VB.TextBox txt_dependente 
            Height          =   285
            Left            =   7200
            MaxLength       =   3
            TabIndex        =   53
            Top             =   5220
            Width           =   615
         End
         Begin VB.TextBox txt_orgao_emissor 
            Height          =   285
            Left            =   7200
            MaxLength       =   10
            TabIndex        =   21
            Top             =   2340
            Width           =   1095
         End
         Begin VB.TextBox txt_serie_carteira_trabalho 
            Height          =   285
            Left            =   7200
            MaxLength       =   10
            TabIndex        =   33
            Top             =   3420
            Width           =   1095
         End
         Begin MSMask.MaskEdBox msk_data_demissao 
            Height          =   300
            Left            =   7200
            TabIndex        =   49
            Top             =   4860
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox msk_data_admissao 
            Height          =   300
            Left            =   1680
            TabIndex        =   47
            Top             =   4860
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox msk_data_nascimento 
            Height          =   300
            Left            =   7200
            TabIndex        =   45
            Top             =   4500
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txt_telefone 
            Height          =   315
            Left            =   7200
            TabIndex        =   17
            Top             =   1980
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   14
            Format          =   "(___) ___-____"
            Mask            =   "(###) ###-####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label9 
            Caption         =   "C.E.P."
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   1980
            Width           =   1455
         End
         Begin VB.Label Label8 
            Caption         =   "Sigla do Estado"
            Height          =   255
            Left            =   5640
            TabIndex        =   12
            Top             =   1620
            Width           =   1455
         End
         Begin VB.Label Label7 
            Caption         =   "Cidade"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   1620
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Código"
            Height          =   255
            Left            =   120
            TabIndex        =   2
            Top             =   180
            Width           =   1455
         End
         Begin VB.Label Label2 
            Caption         =   "Nome"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   540
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "Endereço"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   6
            Top             =   900
            Width           =   1455
         End
         Begin VB.Label Label4 
            Caption         =   "Bairro"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   1260
            Width           =   1455
         End
         Begin VB.Label Label5 
            Caption         =   "C.P.F."
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   2700
            Width           =   1455
         End
         Begin VB.Label Label6 
            Caption         =   "Identidade"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   2340
            Width           =   1455
         End
         Begin VB.Label Label10 
            Caption         =   "Telefone"
            Height          =   255
            Left            =   5640
            TabIndex        =   16
            Top             =   1980
            Width           =   1455
         End
         Begin VB.Label Label30 
            Caption         =   "Cargo"
            Height          =   255
            Index           =   0
            Left            =   5640
            TabIndex        =   24
            Top             =   2700
            Width           =   1455
         End
         Begin VB.Label Label31 
            Caption         =   "Situação"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   3060
            Width           =   1455
         End
         Begin VB.Label Label30 
            Caption         =   "Numero do PIS"
            Height          =   255
            Index           =   1
            Left            =   5640
            TabIndex        =   28
            Top             =   3060
            Width           =   1455
         End
         Begin VB.Label Label30 
            Caption         =   "N. Carteira Trabalho"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   30
            Top             =   3420
            Width           =   1455
         End
         Begin VB.Label Label30 
            Caption         =   "Numero da C.N.H."
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   34
            Top             =   3780
            Width           =   1455
         End
         Begin VB.Label Label30 
            Caption         =   "N. Titulo Eleitor"
            Height          =   255
            Index           =   4
            Left            =   5640
            TabIndex        =   36
            Top             =   3780
            Width           =   1455
         End
         Begin VB.Label Label30 
            Caption         =   "Numero Reservista"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   38
            Top             =   4140
            Width           =   1455
         End
         Begin VB.Label Label30 
            Caption         =   "Sexo"
            Height          =   255
            Index           =   6
            Left            =   5640
            TabIndex        =   40
            Top             =   4140
            Width           =   1455
         End
         Begin VB.Label Label30 
            Caption         =   "Grau de Instrucão"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   42
            Top             =   4500
            Width           =   1455
         End
         Begin VB.Label Label30 
            Caption         =   "Data Nascimento"
            Height          =   255
            Index           =   8
            Left            =   5640
            TabIndex        =   44
            Top             =   4500
            Width           =   1455
         End
         Begin VB.Label Label30 
            Caption         =   "Data de Admissão"
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   46
            Top             =   4860
            Width           =   1455
         End
         Begin VB.Label Label30 
            Caption         =   "Data de Demissão"
            Height          =   255
            Index           =   10
            Left            =   5640
            TabIndex        =   48
            Top             =   4860
            Width           =   1455
         End
         Begin VB.Label Label30 
            Caption         =   "Estado Civil"
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   50
            Top             =   5220
            Width           =   1455
         End
         Begin VB.Label Label30 
            Caption         =   "N. de Dependente"
            Height          =   255
            Index           =   12
            Left            =   5640
            TabIndex        =   52
            Top             =   5220
            Width           =   1455
         End
         Begin VB.Label Label32 
            Caption         =   "Órgão Emissor"
            Height          =   255
            Left            =   5640
            TabIndex        =   20
            Top             =   2340
            Width           =   1455
         End
         Begin VB.Label Label30 
            Caption         =   "Série da Cart. Trab."
            Height          =   255
            Index           =   13
            Left            =   5640
            TabIndex        =   32
            Top             =   3420
            Width           =   1455
         End
      End
      Begin Threed.SSFrame frm_dados_2 
         Height          =   5655
         Left            =   -74880
         TabIndex        =   54
         Top             =   360
         Width           =   9855
         _Version        =   65536
         _ExtentX        =   17383
         _ExtentY        =   9975
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
         Begin VB.TextBox txt_pessoal_nome 
            Height          =   285
            Index           =   0
            Left            =   2100
            MaxLength       =   40
            TabIndex        =   56
            Top             =   180
            Width           =   4815
         End
         Begin VB.TextBox txt_pessoal_bairro 
            Height          =   285
            Index           =   0
            Left            =   2100
            MaxLength       =   30
            TabIndex        =   60
            Top             =   780
            Width           =   3735
         End
         Begin VB.TextBox txt_pessoal_endereco 
            Height          =   285
            Index           =   0
            Left            =   2100
            MaxLength       =   40
            TabIndex        =   58
            Top             =   480
            Width           =   4815
         End
         Begin VB.TextBox txt_pessoal_cidade 
            Height          =   285
            Index           =   0
            Left            =   2100
            MaxLength       =   20
            TabIndex        =   62
            Top             =   1080
            Width           =   2535
         End
         Begin VB.TextBox txt_pessoal_grau_relacionamento 
            Height          =   285
            Index           =   0
            Left            =   2100
            MaxLength       =   40
            TabIndex        =   66
            Top             =   1380
            Width           =   4815
         End
         Begin VB.TextBox txt_pessoal_nome 
            Height          =   285
            Index           =   1
            Left            =   2100
            MaxLength       =   40
            TabIndex        =   68
            Top             =   1920
            Width           =   4815
         End
         Begin VB.TextBox txt_pessoal_bairro 
            Height          =   285
            Index           =   1
            Left            =   2100
            MaxLength       =   30
            TabIndex        =   72
            Top             =   2520
            Width           =   3735
         End
         Begin VB.TextBox txt_pessoal_endereco 
            Height          =   285
            Index           =   1
            Left            =   2100
            MaxLength       =   40
            TabIndex        =   70
            Top             =   2220
            Width           =   4815
         End
         Begin VB.TextBox txt_pessoal_cidade 
            Height          =   285
            Index           =   1
            Left            =   2100
            MaxLength       =   20
            TabIndex        =   74
            Top             =   2820
            Width           =   2535
         End
         Begin VB.TextBox txt_pessoal_grau_relacionamento 
            Height          =   285
            Index           =   1
            Left            =   2100
            MaxLength       =   40
            TabIndex        =   78
            Top             =   3120
            Width           =   4815
         End
         Begin VB.TextBox txt_pessoal_nome 
            Height          =   285
            Index           =   2
            Left            =   2100
            MaxLength       =   40
            TabIndex        =   80
            Top             =   3600
            Width           =   4815
         End
         Begin VB.TextBox txt_pessoal_bairro 
            Height          =   285
            Index           =   2
            Left            =   2100
            MaxLength       =   30
            TabIndex        =   84
            Top             =   4200
            Width           =   3735
         End
         Begin VB.TextBox txt_pessoal_endereco 
            Height          =   285
            Index           =   2
            Left            =   2100
            MaxLength       =   40
            TabIndex        =   82
            Top             =   3900
            Width           =   4815
         End
         Begin VB.TextBox txt_pessoal_cidade 
            Height          =   285
            Index           =   2
            Left            =   2100
            MaxLength       =   20
            TabIndex        =   86
            Top             =   4500
            Width           =   2535
         End
         Begin VB.TextBox txt_pessoal_grau_relacionamento 
            Height          =   285
            Index           =   2
            Left            =   2100
            MaxLength       =   40
            TabIndex        =   90
            Top             =   4800
            Width           =   4815
         End
         Begin MSMask.MaskEdBox txt_pessoal_telefone 
            Height          =   315
            Index           =   2
            Left            =   8340
            TabIndex        =   88
            Top             =   4500
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   14
            Format          =   "0000000000"
            Mask            =   "(###) ###-####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txt_pessoal_telefone 
            Height          =   315
            Index           =   1
            Left            =   8340
            TabIndex        =   76
            Top             =   2820
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   14
            Format          =   "0000000000"
            Mask            =   "(###) ###-####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txt_pessoal_telefone 
            Height          =   315
            Index           =   0
            Left            =   8340
            TabIndex        =   64
            Top             =   1080
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   14
            Format          =   "0000000000"
            Mask            =   "(###) ###-####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label13 
            Caption         =   "Nome"
            Height          =   255
            Left            =   120
            TabIndex        =   55
            Top             =   180
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Telefone"
            Height          =   255
            Left            =   6840
            TabIndex        =   63
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label Label14 
            Caption         =   "Bairro"
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Top             =   780
            Width           =   1455
         End
         Begin VB.Label Label15 
            Caption         =   "Endereço"
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label16 
            Caption         =   "Cidade"
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label Label17 
            Caption         =   "Grau de Relacionamento"
            Height          =   255
            Left            =   120
            TabIndex        =   65
            Top             =   1380
            Width           =   1875
         End
         Begin VB.Label Label18 
            Caption         =   "Nome"
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   1920
            Width           =   1455
         End
         Begin VB.Label Label19 
            Caption         =   "Telefone"
            Height          =   255
            Left            =   6840
            TabIndex        =   75
            Top             =   2820
            Width           =   1455
         End
         Begin VB.Label Label20 
            Caption         =   "Bairro"
            Height          =   255
            Left            =   120
            TabIndex        =   71
            Top             =   2520
            Width           =   1455
         End
         Begin VB.Label Label21 
            Caption         =   "Endereço"
            Height          =   255
            Left            =   120
            TabIndex        =   69
            Top             =   2220
            Width           =   1455
         End
         Begin VB.Label Label22 
            Caption         =   "Cidade"
            Height          =   255
            Left            =   120
            TabIndex        =   73
            Top             =   2820
            Width           =   1455
         End
         Begin VB.Label Label23 
            Caption         =   "Grau de Relacionamento"
            Height          =   255
            Left            =   120
            TabIndex        =   77
            Top             =   3120
            Width           =   1875
         End
         Begin VB.Label Label24 
            Caption         =   "Nome"
            Height          =   255
            Left            =   120
            TabIndex        =   79
            Top             =   3600
            Width           =   1455
         End
         Begin VB.Label Label25 
            Caption         =   "Telefone"
            Height          =   255
            Left            =   6840
            TabIndex        =   87
            Top             =   4500
            Width           =   1455
         End
         Begin VB.Label Label26 
            Caption         =   "Bairro"
            Height          =   255
            Left            =   120
            TabIndex        =   83
            Top             =   4200
            Width           =   1455
         End
         Begin VB.Label Label27 
            Caption         =   "Endereço"
            Height          =   255
            Left            =   120
            TabIndex        =   81
            Top             =   3900
            Width           =   1455
         End
         Begin VB.Label Label28 
            Caption         =   "Cidade"
            Height          =   255
            Left            =   120
            TabIndex        =   85
            Top             =   4500
            Width           =   1455
         End
         Begin VB.Label Label29 
            Caption         =   "Grau de Relacionamento"
            Height          =   255
            Left            =   120
            TabIndex        =   89
            Top             =   4800
            Width           =   1875
         End
      End
      Begin Threed.SSFrame frm_dados_3 
         Height          =   5655
         Left            =   -74880
         TabIndex        =   91
         Top             =   360
         Width           =   9855
         _Version        =   65536
         _ExtentX        =   17383
         _ExtentY        =   9975
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
         Begin VB.TextBox txt_senha 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1680
            MaxLength       =   10
            PasswordChar    =   "*"
            TabIndex        =   119
            Top             =   2700
            Width           =   1035
         End
         Begin VB.TextBox txt_salario_base 
            Height          =   285
            Left            =   1680
            MaxLength       =   10
            TabIndex        =   93
            Top             =   180
            Width           =   1095
         End
         Begin VB.CheckBox chk_adicional_periculosidade 
            Height          =   255
            Left            =   1680
            TabIndex        =   95
            Top             =   540
            Width           =   315
         End
         Begin VB.CheckBox chk_adicional_noturno 
            Height          =   255
            Left            =   1680
            TabIndex        =   97
            Top             =   900
            Width           =   315
         End
         Begin VB.CheckBox chk_cesta_basica 
            Height          =   255
            Left            =   1680
            TabIndex        =   99
            Top             =   1260
            Width           =   315
         End
         Begin VB.CheckBox chk_vale_transporte 
            Height          =   255
            Left            =   1680
            TabIndex        =   101
            Top             =   1620
            Width           =   315
         End
         Begin VB.CheckBox chk_seguro_vida 
            Height          =   255
            Left            =   1680
            TabIndex        =   103
            Top             =   1980
            Width           =   315
         End
         Begin VB.ComboBox cbo_periodo 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   105
            Top             =   2340
            Width           =   495
         End
         Begin VB.Label Label33 
            Caption         =   "Senha"
            Height          =   255
            Left            =   120
            TabIndex        =   120
            Top             =   2700
            Width           =   1515
         End
         Begin VB.Label Label3 
            Caption         =   "Salário base"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   92
            Top             =   180
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "Ad. periculosidade"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   94
            Top             =   540
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "Ad. noturno"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   96
            Top             =   900
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "Cesta básica"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   98
            Top             =   1260
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "Vale transporte"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   100
            Top             =   1620
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "Seguro de vida"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   102
            Top             =   1980
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "&Período"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   104
            Top             =   2340
            Width           =   1515
         End
      End
      Begin VB.Label Label12 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   15
         Left            =   0
         TabIndex        =   118
         Top             =   0
         Width           =   15
      End
   End
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   7980
      TabIndex        =   113
      Top             =   6120
      Width           =   2175
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "cad_funcionario.frx":74C6
         Style           =   1  'Graphical
         TabIndex        =   116
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "cad_funcionario.frx":8A48
         Style           =   1  'Graphical
         TabIndex        =   115
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "cad_funcionario.frx":9EBA
         Style           =   1  'Graphical
         TabIndex        =   117
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "cad_funcionario.frx":B3B4
         Style           =   1  'Graphical
         TabIndex        =   114
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   9360
      Picture         =   "cad_funcionario.frx":C8AE
      Style           =   1  'Graphical
      TabIndex        =   107
      ToolTipText     =   "Cancela o registro atual."
      Top             =   6240
      Width           =   795
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   8460
      Picture         =   "cad_funcionario.frx":DDA8
      Style           =   1  'Graphical
      TabIndex        =   106
      ToolTipText     =   "Confirma o registro atual."
      Top             =   6240
      Width           =   795
   End
End
Attribute VB_Name = "cadastro_funcionario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lFlagCadastro As Integer
Dim lOpcao As Integer
Dim lCodigo As Integer
Dim MyAppID
Private Funcionario As New cFuncionario
Private Sub AtualTabe()
    lCodigo = Val(txt_codigo.Text)
    Funcionario.Empresa = g_empresa
    Funcionario.Codigo = Val(txt_codigo.Text)
    Funcionario.Nome = txt_nome.Text
    Funcionario.Endereco = txt_endereco.Text
    Funcionario.Bairro = txt_bairro.Text
    Funcionario.Cidade = txt_cidade.Text
    Funcionario.UF = txt_uf.Text
    Funcionario.CEP = txt_cep.Text
    Funcionario.CPF = txt_cpf.Text
    Funcionario.Identidade = txt_identidade.Text
    Funcionario.OrgaoEmissor = txt_orgao_emissor.Text
    Funcionario.Telefone = txt_telefone.Text
    Funcionario.Cargo = txt_cargo.Text
    Funcionario.Situacao = txt_situacao.Text
    Funcionario.PIS = txt_pis.Text
    Funcionario.CarteiraTrabalho = txt_carteira_trabalho.Text
    Funcionario.SerieCarteiraTrabalho = txt_serie_carteira_trabalho.Text
    Funcionario.CNH = Format(txt_cnh.Text, "")
    Funcionario.TituloEleitor = txt_titulo_eleitor.Text
    Funcionario.NumeroReservista = txt_reservista.Text
    Funcionario.Sexo = txt_sexo.Text
    Funcionario.GrauInstrucao = txt_grau_instrucao.Text
    If IsDate(msk_data_nascimento) Then
        Funcionario.DataNascimento = Format(msk_data_nascimento, "dd/mm/yyyy")
    Else
        Funcionario.DataNascimento = "00:00:00"
    End If
    If IsDate(msk_data_admissao) Then
        Funcionario.DataAdmissao = Format(msk_data_admissao, "dd/mm/yyyy")
    Else
        Funcionario.DataAdmissao = "00:00:00"
    End If
    If IsDate(msk_data_demissao) Then
        Funcionario.DataDemissao = Format(msk_data_demissao, "dd/mm/yyyy")
    Else
        Funcionario.DataDemissao = "00:00:00"
    End If
    Funcionario.EstadoCivil = txt_estado_civil.Text
    Funcionario.Dependente = Val(txt_dependente.Text)
    Funcionario.PessoalNome1 = txt_pessoal_nome(0)
    Funcionario.PessoalEndereco1 = txt_pessoal_endereco(0)
    Funcionario.PessoalBairro1 = txt_pessoal_bairro(0)
    Funcionario.PessoalCidade1 = txt_pessoal_cidade(0)
    Funcionario.PessoalTelefone1 = txt_pessoal_telefone(0)
    Funcionario.PessoalGrauRelacionamento1 = txt_pessoal_grau_relacionamento(0)
    Funcionario.PessoalNome2 = txt_pessoal_nome(1)
    Funcionario.PessoalEndereco2 = txt_pessoal_endereco(1)
    Funcionario.PessoalBairro2 = txt_pessoal_bairro(1)
    Funcionario.PessoalCidade2 = txt_pessoal_cidade(1)
    Funcionario.PessoalTelefone2 = txt_pessoal_telefone(1)
    Funcionario.PessoalGrauRelacionamento2 = txt_pessoal_grau_relacionamento(1)
    Funcionario.PessoalNome3 = txt_pessoal_nome(2)
    Funcionario.PessoalEndereco3 = txt_pessoal_endereco(2)
    Funcionario.PessoalBairro3 = txt_pessoal_bairro(2)
    Funcionario.PessoalCidade3 = txt_pessoal_cidade(2)
    Funcionario.PessoalTelefone3 = txt_pessoal_telefone(2)
    Funcionario.PessoalGrauRelacionamento3 = txt_pessoal_grau_relacionamento(2)
    Funcionario.SalarioBase = fValidaValor2(txt_salario_base.Text)
    Funcionario.AdicionalPericulosidade = chk_adicional_periculosidade.Value
    Funcionario.AdicionalNoturno = chk_adicional_noturno.Value
    Funcionario.CestaBasica = chk_cesta_basica.Value
    Funcionario.ValeTransporte = chk_vale_transporte.Value
    Funcionario.SeguroVida = chk_seguro_vida.Value
    Funcionario.Periodo = Val(cbo_periodo)
    Funcionario.Senha = Kriptografa(txt_senha.Text)
End Sub
Private Sub AtualTela()
    tab_dados.Tab = 0
    lCodigo = Val(Funcionario.Codigo)
    txt_codigo.Text = Funcionario.Codigo
    txt_nome.Text = Funcionario.Nome
    txt_endereco.Text = Funcionario.Endereco
    txt_bairro.Text = Funcionario.Bairro
    txt_cidade.Text = Funcionario.Cidade
    txt_uf.Text = Funcionario.UF
    txt_cep.Text = Funcionario.CEP
    txt_cpf.Text = Funcionario.CPF
    txt_identidade.Text = Funcionario.Identidade
    txt_orgao_emissor.Text = Format(Funcionario.OrgaoEmissor, "")
    If Len(Trim(Funcionario.Telefone)) > 0 Then
        txt_telefone.Text = Format(Funcionario.Telefone, "(000) 000-0000")
    Else
        txt_telefone.Text = ""
    End If
    txt_cargo.Text = Funcionario.Cargo
    txt_situacao.Text = Funcionario.Situacao
    txt_pis.Text = Funcionario.PIS
    txt_carteira_trabalho.Text = Funcionario.CarteiraTrabalho
    txt_serie_carteira_trabalho.Text = Format(Funcionario.SerieCarteiraTrabalho, "00000")
    txt_cnh.Text = Funcionario.CNH
    txt_titulo_eleitor.Text = Funcionario.TituloEleitor
    txt_reservista.Text = Funcionario.NumeroReservista
    txt_sexo.Text = Funcionario.Sexo
    txt_grau_instrucao.Text = Funcionario.GrauInstrucao
    If Funcionario.DataNascimento <> "00:00:00" Then
        msk_data_nascimento = Format(Funcionario.DataNascimento, "dd/mm/yyyy")
    Else
        msk_data_nascimento = "__/__/____"
    End If
    If Funcionario.DataAdmissao <> "00:00:00" Then
        msk_data_admissao = Format(Funcionario.DataAdmissao, "dd/mm/yyyy")
    Else
        msk_data_admissao = "__/__/____"
    End If
    If Funcionario.DataDemissao <> "00:00:00" Then
        msk_data_demissao = Format(Funcionario.DataDemissao, "dd/mm/yyyy")
    Else
        msk_data_demissao = "__/__/____"
    End If
    txt_estado_civil.Text = Funcionario.EstadoCivil
    txt_dependente.Text = Funcionario.Dependente
    txt_pessoal_nome(0) = Funcionario.PessoalNome1
    txt_pessoal_endereco(0) = Funcionario.PessoalEndereco1
    txt_pessoal_bairro(0) = Funcionario.PessoalBairro1
    txt_pessoal_cidade(0) = Funcionario.PessoalCidade1
    If Len(Trim(Funcionario.PessoalTelefone1)) > 0 Then
        txt_pessoal_telefone(0) = Format(Funcionario.PessoalTelefone1, "(000) 000-0000")
    Else
        txt_pessoal_telefone(0) = ""
    End If
    txt_pessoal_grau_relacionamento(0) = Funcionario.PessoalGrauRelacionamento1
    txt_pessoal_nome(1) = Funcionario.PessoalNome2
    txt_pessoal_endereco(1) = Funcionario.PessoalEndereco2
    txt_pessoal_bairro(1) = Funcionario.PessoalBairro2
    txt_pessoal_cidade(1) = Funcionario.PessoalCidade2
    If Len(Trim(Funcionario.PessoalTelefone2)) > 0 Then
        txt_pessoal_telefone(1) = Format(Funcionario.PessoalTelefone2, "(000) 000-0000")
    Else
        txt_pessoal_telefone(1) = ""
    End If
    txt_pessoal_grau_relacionamento(1) = Funcionario.PessoalGrauRelacionamento2
    txt_pessoal_nome(2) = Funcionario.PessoalNome3
    txt_pessoal_endereco(2) = Funcionario.PessoalEndereco3
    txt_pessoal_bairro(2) = Funcionario.PessoalBairro3
    txt_pessoal_cidade(2) = Funcionario.PessoalCidade3
    If Len(Trim(Funcionario.PessoalTelefone3)) > 0 Then
        txt_pessoal_telefone(2) = Format(Funcionario.PessoalTelefone3, "(000) 000-0000")
    Else
        txt_pessoal_telefone(2) = ""
    End If
    txt_pessoal_grau_relacionamento(2) = Funcionario.PessoalGrauRelacionamento3
    txt_salario_base.Text = Format(Funcionario.SalarioBase, "###,##0.00")
    If Funcionario.AdicionalPericulosidade Then
        chk_adicional_periculosidade.Value = 1
    Else
        chk_adicional_periculosidade.Value = 0
    End If
    If Funcionario.AdicionalNoturno Then
        chk_adicional_noturno.Value = 1
    Else
        chk_adicional_noturno.Value = 0
    End If
    If Funcionario.CestaBasica Then
        chk_cesta_basica.Value = 1
    Else
        chk_cesta_basica.Value = 0
    End If
    If Funcionario.ValeTransporte Then
        chk_vale_transporte.Value = 1
    Else
        chk_vale_transporte.Value = 0
    End If
    If Funcionario.SeguroVida Then
        chk_seguro_vida.Value = 1
    Else
        chk_seguro_vida.Value = 0
    End If
    cbo_periodo.ListIndex = Funcionario.Periodo
    txt_senha.Text = DesKriptografa(Funcionario.Senha)
    frm_dados_1.Enabled = False
    frm_dados_2.Enabled = False
    frm_dados_3.Enabled = False
End Sub
Private Sub Finaliza()
    Set Funcionario = Nothing
    frm_cadastro.Show
End Sub
Private Sub Inclui()
    lOpcao = 1
    DesativaBotoes
    cmd_novo.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    txt_codigo.Text = 1
    If Funcionario.LocalizarUltimo(g_empresa) Then
        txt_codigo.Text = Funcionario.Codigo + 1
    End If
End Sub
Private Sub cbo_periodo_GotFocus()
    SendMessageLong cbo_periodo.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_periodo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_senha.SetFocus
    End If
End Sub

Private Sub chk_adicional_noturno_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        chk_cesta_basica.SetFocus
    End If
End Sub
Private Sub chk_adicional_periculosidade_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        chk_adicional_noturno.SetFocus
    End If
End Sub
Private Sub chk_cesta_basica_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        chk_vale_transporte.SetFocus
    End If
End Sub
Private Sub chk_seguro_vida_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_periodo.SetFocus
    End If
End Sub
Private Sub chk_vale_transporte_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        chk_seguro_vida.SetFocus
    End If
End Sub
Private Sub cmd_alterar_Click()
    lOpcao = 2
    DesativaBotoes
    cmd_alterar.Visible = True
    cmd_alterar.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    frm_dados_1.Enabled = True
    frm_dados_2.Enabled = True
    frm_dados_3.Enabled = True
    txt_codigo.Enabled = False
    txt_nome.SetFocus
End Sub
Private Sub cmd_anterior_Click()
    If Funcionario.LocalizarAnterior Then
        AtualTela
    Else
        MsgBox "Início de Arquivo.", vbInformation, "Atenção!"
        cmd_proximo.SetFocus
    End If
End Sub
Private Sub cmd_cancelar_Click()
    LimpaTela
    lOpcao = 0
    If Funcionario.LocalizarCodigo(g_empresa, lCodigo) Then
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
    Dim i As Integer
    txt_codigo = ""
    txt_nome = ""
    txt_endereco = ""
    txt_bairro = ""
    txt_cidade = ""
    txt_uf = ""
    txt_cep = ""
    txt_cpf = ""
    txt_identidade = ""
    txt_orgao_emissor = ""
    txt_telefone = ""
    txt_cargo = ""
    txt_situacao = ""
    txt_pis = ""
    txt_carteira_trabalho = ""
    txt_serie_carteira_trabalho = ""
    txt_cnh = ""
    txt_titulo_eleitor = ""
    txt_reservista = ""
    txt_sexo = ""
    txt_grau_instrucao = ""
    msk_data_nascimento = "__/__/____"
    msk_data_admissao = "__/__/____"
    msk_data_demissao = "__/__/____"
    txt_estado_civil = ""
    txt_dependente = ""
    For i = 0 To 2
        txt_pessoal_nome(i) = ""
        txt_pessoal_endereco(i) = ""
        txt_pessoal_bairro(i) = ""
        txt_pessoal_cidade(i) = ""
        txt_pessoal_telefone(i) = ""
        txt_pessoal_grau_relacionamento(i) = ""
        txt_salario_base = ""
    Next
End Sub
Private Sub PreencheCboPeriodo()
    cbo_periodo.Clear
    cbo_periodo.AddItem 0
    cbo_periodo.ItemData(cbo_periodo.NewIndex) = 0
    cbo_periodo.AddItem 1
    cbo_periodo.ItemData(cbo_periodo.NewIndex) = 1
    cbo_periodo.AddItem 2
    cbo_periodo.ItemData(cbo_periodo.NewIndex) = 2
    cbo_periodo.AddItem 3
    cbo_periodo.ItemData(cbo_periodo.NewIndex) = 3
    cbo_periodo.AddItem 4
    cbo_periodo.ItemData(cbo_periodo.NewIndex) = 4
    cbo_periodo.AddItem 5
    cbo_periodo.ItemData(cbo_periodo.NewIndex) = 5
End Sub
Private Sub cmd_excluir_Click()
    If txt_codigo.Text <> "" Then
        If (MsgBox("Deseja Realmente Excluir Este Registro?", 4 + 32 + 256, "Exclusão de Registro!")) = 6 Then
            If Funcionario.Excluir(g_empresa, txt_codigo.Text) Then
                LimpaTela
                If Funcionario.LocalizarUltimo(g_empresa) Then
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
Private Sub cmd_fechar_Click()
    SendKeys "%{F4}"
    AppActivate MyAppID ' Activate Microsoft
End Sub
Private Sub cmd_novo_Click()
    LimpaTela
    Inclui
    frm_dados_1.Enabled = True
    frm_dados_2.Enabled = True
    frm_dados_3.Enabled = True
    txt_nome.SetFocus
End Sub
Private Sub cmd_ok_Click()
    On Error GoTo FileError
    If ValidaCampos Then
        AtivaBotoes
        If lOpcao = 1 Then
            AtualTabe
            If Funcionario.Incluir Then
                lCodigo = txt_codigo.Text
            Else
                MsgBox "Não foi possível incluir este registro!", vbInformation, "Erro de Verificação!"
            End If
        ElseIf lOpcao = 2 Then
            AtualTabe
            If Not Funcionario.Alterar(g_empresa, lCodigo) Then
                MsgBox "Não foi possível alterar este registro!", vbInformation, "Erro de Verificação!"
            End If
        End If
        lOpcao = 0
        Call Funcionario.LocalizarCodigo(g_empresa, lCodigo)
        AtualTela
        cmd_novo.SetFocus
    End If
    Exit Sub
FileError:
    MsgBox Error
    'ErroArquivo tbl_funcionario.Name, "Funcionárioo"
    Exit Sub
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If Not Val(txt_codigo) > 0 Then
        MsgBox "Informe o código do funcionário.", vbInformation, "Atenção!"
        tab_dados.Tab = 0
        txt_codigo.SetFocus
    ElseIf Not txt_nome > "" Then
        MsgBox "Informe o nome do funcionário.", vbInformation, "Atenção!"
        tab_dados.Tab = 0
        txt_nome.SetFocus
    'ElseIf Not txt_endereco > "" Then
    '    MsgBox "Informe o Endereco.", vbInformation, "Atenção!"
    '    tab_dados.Tab = 0
    '    txt_endereco.SetFocus
    'ElseIf Not txt_bairro > "" Then
    '    MsgBox "Informe o Bairro.", vbInformation, "Atenção!"
    '    tab_dados.Tab = 0
    '    txt_bairro.SetFocus
    'ElseIf Not txt_cidade > "" Then
    '    MsgBox "Informe a Cidade.", vbInformation, "Atenção!"
    '    tab_dados.Tab = 0
    '    txt_cidade.SetFocus
    'ElseIf Not txt_uf > "" Then
    '    MsgBox "Informe a Sigla do Estado.", vbInformation, "Atenção!"
    '    tab_dados.Tab = 0
    '    txt_uf.SetFocus
    'ElseIf Not Val(txt_cep) > 0 Then
    '    MsgBox "Informe o CEP.", vbInformation, "Atenção!"
    '    tab_dados.Tab = 0
    '    txt_cep.SetFocus
    'ElseIf Not txt_cpf > "" Then
    '    MsgBox "Informe o Número do CPF.", vbInformation, "Atenção!"
    '    tab_dados.Tab = 0
    '    txt_cpf.SetFocus
    'ElseIf Not txt_identidade > "" Then
    '    MsgBox "Informe o Número da Identidade.", vbInformation, "Atenção!"
    '    tab_dados.Tab = 0
    '    txt_identidade.SetFocus
    'ElseIf Not txt_telefone > "" Then
    '    MsgBox "Informe o Número do Telefone.", vbInformation, "Atenção!"
    '    tab_dados.Tab = 0
    '    txt_telefone.SetFocus
    ElseIf txt_sexo <> "M" And txt_sexo <> "F" Then
        MsgBox "Informe o Sexo - 'M' ou 'F'.", vbInformation, "Atenção!"
        tab_dados.Tab = 0
        txt_sexo.SetFocus
    ElseIf Not txt_cargo > "" Then
        MsgBox "Informe o cargo do funcionário.", vbInformation, "Atenção!"
        tab_dados.Tab = 0
        txt_cargo.SetFocus
    ElseIf txt_situacao <> "A" And txt_situacao <> "I" Then
        MsgBox "Informe Ativo ou Inativo.", vbInformation, "Atenção!"
        tab_dados.Tab = 0
        txt_situacao.SetFocus
    ElseIf txt_senha = "" Then
        MsgBox "Informe a senha de acesso do funcionário.", vbInformation, "Atenção!"
        txt_senha.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_pesquisa_Click()
    consulta_funcionario.Show 1
    If Len(g_string) > 0 Then
        lCodigo = RetiraGString(1)
        If Funcionario.LocalizarCodigo(g_empresa, lCodigo) Then
            AtualTela
        End If
    End If
End Sub
Private Sub cmd_primeiro_Click()
    If Funcionario.LocalizarPrimeiro Then
        AtualTela
        cmd_proximo.SetFocus
    Else
        MsgBox "Não há registro.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    If Funcionario.LocalizarProximo Then
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
    If Funcionario.LocalizarUltimo(g_empresa) Then
        AtualTela
        cmd_anterior.SetFocus
    Else
        MsgBox "Não há registro.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub Z_KBG_ConverteFuncionarioBancoDados()
    Dim old_sgp As Database
    Dim old_funcionario As Table
    
    Set old_sgp = OpenDatabase("d:\backup\vb4\sgp\SGP_DATA.MDB")
    Set old_funcionario = old_sgp.OpenTable("Funcionario")
    old_funcionario.Index = "id_codigo"
    
    old_funcionario.Seek "=", 4, 12
    If Not old_funcionario.NoMatch Then
    
        With tbl_funcionario
            .AddNew
            !Empresa = old_funcionario!Empresa
            !Codigo = old_funcionario!Codigo
            !Nome = old_funcionario!Nome
            !Endereco = old_funcionario!Endereco
            !Bairro = old_funcionario!Bairro
            !Cidade = old_funcionario!Cidade
            !UF = old_funcionario!UF
            !CEP = old_funcionario!CEP
            !CPF = old_funcionario!CPF
            !Identidade = old_funcionario!Identidade
            !Telefone = old_funcionario!Telefone
            !Cargo = old_funcionario!Cargo
            !Situacao = old_funcionario!Situacao
            !PIS = old_funcionario!PIS
            ![Carteira de Trabalho] = old_funcionario![Carteira de Trabalho]
            !CNH = old_funcionario!CNH
            ![Titulo de Eleitor] = old_funcionario![Titulo de Eleitor]
            ![Numero da Reservista] = old_funcionario![Numero da Reservista]
            !Sexo = old_funcionario!Sexo
            ![Grau de Instrucao] = old_funcionario![Grau de Instrucao]
'            If IsDate(msk_data_nascimento) Then
'            Else
'                ![Data de Nascimento] = Null
'            End If
            ![Data de Nascimento] = old_funcionario![Data de Nascimento]
            ![Data de Admissao] = old_funcionario![Data de Admissao]
            ![Data de Demissao] = old_funcionario![Data de Demissao]
            ![Estado Civil] = old_funcionario![Estado Civil]
            !Dependente = old_funcionario!Dependente
            ![Pessoal Nome 1] = old_funcionario![Pessoal Nome 1]
            ![Pessoal Endereco 1] = old_funcionario![Pessoal Endereco 1]
            ![Pessoal Bairro 1] = old_funcionario![Pessoal Bairro 1]
            ![Pessoal Cidade 1] = old_funcionario![Pessoal Cidade 1]
            ![Pessoal Telefone 1] = old_funcionario![Pessoal Telefone 1]
            ![Pessoal Grau Relacionamento 1] = old_funcionario![Pessoal Grau Relacionamento 1]
            ![Pessoal Nome 2] = old_funcionario![Pessoal Nome 2]
            ![Pessoal Endereco 2] = old_funcionario![Pessoal Endereco 2]
            ![Pessoal Bairro 2] = old_funcionario![Pessoal Bairro 2]
            ![Pessoal Cidade 2] = old_funcionario![Pessoal Cidade 2]
            ![Pessoal Telefone 2] = old_funcionario![Pessoal Telefone 2]
            ![Pessoal Grau Relacionamento 2] = old_funcionario![Pessoal Grau Relacionamento 2]
            ![Pessoal Nome 3] = old_funcionario![Pessoal Nome 3]
            ![Pessoal Endereco 3] = old_funcionario![Pessoal Endereco 3]
            ![Pessoal Bairro 3] = old_funcionario![Pessoal Bairro 3]
            ![Pessoal Cidade 3] = old_funcionario![Pessoal Cidade 3]
            ![Pessoal Telefone 3] = old_funcionario![Pessoal Telefone 3]
            ![Pessoal Grau Relacionamento 3] = old_funcionario![Pessoal Grau Relacionamento 3]
            .Update
        End With
    old_funcionario.Close
    old_sgp.Close
    End If
End Sub
Private Sub Form_Activate()
    If lFlagCadastro = 0 Then
        PreencheCboPeriodo
        DesativaBotoes
        If Funcionario.LocalizarUltimo(g_empresa) Then
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
    Screen.MousePointer = 1
End Sub
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
    CentraForm Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub

Private Sub msk_data_admissao_GotFocus()
'    msk_data_admissao.SelLength = Len(msk_data_admissao)
End Sub
Private Sub msk_data_admissao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{Tab}"
    End If
End Sub
Private Sub msk_data_demissao_GotFocus()
'     msk_data_demissao.SelLength = Len(msk_data_demissao)
End Sub
Private Sub msk_data_demissao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{Tab}"
    End If
End Sub
Private Sub msk_data_nascimento_GotFocus()
'    msk_data_nascimento.SelLength = Len(msk_data_nascimento)
End Sub
Private Sub msk_data_nascimento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{Tab}"
    End If
End Sub
Private Sub txt_bairro_GotFocus()
    txt_bairro.SelLength = Len(txt_bairro)
End Sub
Private Sub txt_bairro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{Tab}"
    End If
End Sub
Private Sub txt_cargo_GotFocus()
    txt_cargo.SelLength = Len(txt_cargo)
End Sub
Private Sub txt_cargo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{Tab}"
    End If
End Sub
Private Sub txt_carteira_trabalho_GotFocus()
    txt_carteira_trabalho.SelLength = Len(txt_carteira_trabalho)
End Sub
Private Sub txt_carteira_trabalho_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{Tab}"
    End If
End Sub
Private Sub txt_cep_GotFocus()
    txt_cep.SelLength = Len(txt_cep)
End Sub
Private Sub txt_cep_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{Tab}"
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_cidade_GotFocus()
    txt_cidade.SelLength = Len(txt_cidade)
End Sub
Private Sub txt_cidade_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{Tab}"
    End If
End Sub
Private Sub txt_cnh_GotFocus()
    txt_cnh.SelLength = Len(txt_cnh)
End Sub
Private Sub txt_cnh_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{Tab}"
    End If
End Sub
Private Sub txt_codigo_GotFocus()
    txt_codigo.SelLength = Len(txt_codigo)
End Sub
Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{Tab}"
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_codigo_LostFocus()
    If lOpcao = 1 And txt_codigo.Text <> "" Then
        If Funcionario.LocalizarCodigo(g_empresa, Val(txt_codigo.Text)) Then
            MsgBox "Já existe funcionário cadastrado com este código." & Chr(10) & Chr(10) & "Mude o código informado.", vbInformation, "Duplicidade de Registro!"
            txt_codigo.SetFocus
            Exit Sub
        End If
    End If
End Sub
Private Sub txt_cpf_GotFocus()
    txt_cpf.SelLength = Len(txt_cpf)
End Sub
Private Sub txt_cpf_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{Tab}"
    End If
End Sub
Private Sub txt_dependente_GotFocus()
    txt_dependente.SelLength = Len(txt_dependente)
End Sub
Private Sub txt_dependente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        tab_dados.Tab = 1
        txt_pessoal_nome(0).SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_dependente_LostFocus()
    txt_dependente = Format(txt_dependente, "#0")
End Sub
Private Sub txt_endereco_GotFocus()
    txt_endereco.SelLength = Len(txt_endereco)
End Sub
Private Sub txt_endereco_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{Tab}"
    End If
End Sub
Private Sub txt_estado_civil_GotFocus()
    txt_estado_civil.SelLength = Len(txt_estado_civil)
End Sub
Private Sub txt_estado_civil_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{Tab}"
    End If
End Sub
Private Sub txt_grau_instrucao_GotFocus()
    txt_grau_instrucao.SelLength = Len(txt_grau_instrucao)
End Sub
Private Sub txt_grau_instrucao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{Tab}"
    End If
End Sub
Private Sub txt_identidade_GotFocus()
    txt_identidade.SelLength = Len(txt_identidade)
End Sub
Private Sub txt_identidade_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{Tab}"
    End If
End Sub
Private Sub txt_nome_GotFocus()
    txt_nome.SelLength = Len(txt_nome)
End Sub
Private Sub txt_nome_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{Tab}"
    End If
End Sub
Private Sub txt_nome_LostFocus()
    If lOpcao = 1 And txt_nome <> "" Then
        If Funcionario.LocalizarNome(txt_nome.Text) Then
            If (MsgBox("Existe outro funcionário cadastrado com o mesmo nome." & Chr(10) & Chr(10) & "Deseja cadastrar assim mesmo?", 4 + 32 + 256, "Duplicidade de Registro!")) = 7 Then
                txt_nome.SetFocus
            End If
        End If
        'tbl_funcionario.Index = "id_nome"
        'tbl_funcionario.Seek ">=", g_empresa, txt_nome, 0
        'If Not tbl_funcionario.NoMatch Then
        '    If tbl_funcionario!Empresa = g_empresa Then
        '        If UCase(tbl_funcionario!Nome) = UCase(txt_nome) Then
        '            If (MsgBox("Já existe funcionário cadastrado com este nome." & Chr(10) & Chr(10) & "Deseja cadastrar assim mesmo?", 4 + 32 + 256, "Duplicidade de Registro!")) = 7 Then
        '                txt_nome.SetFocus
        '            End If
        '        End If
        '    End If
        'End If
        'tbl_funcionario.Index = "id_codigo"
    End If
End Sub
Private Sub txt_orgao_emissor_GotFocus()
    txt_orgao_emissor.SelLength = Len(txt_orgao_emissor)
End Sub
Private Sub txt_orgao_emissor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{Tab}"
    End If
End Sub
Private Sub txt_pessoal_bairro_GotFocus(Index As Integer)
    txt_pessoal_bairro(Index).SelLength = Len(txt_pessoal_bairro(Index))
End Sub
Private Sub txt_pessoal_bairro_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{Tab}"
    End If
End Sub
Private Sub txt_pessoal_cidade_GotFocus(Index As Integer)
    txt_pessoal_cidade(Index).SelLength = Len(txt_pessoal_cidade(Index))
End Sub
Private Sub txt_pessoal_cidade_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{Tab}"
    End If
End Sub
Private Sub txt_pessoal_endereco_GotFocus(Index As Integer)
    txt_pessoal_endereco(Index).SelLength = Len(txt_pessoal_endereco(Index))
End Sub
Private Sub txt_pessoal_endereco_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{Tab}"
    End If
End Sub
Private Sub txt_pessoal_grau_relacionamento_GotFocus(Index As Integer)
    txt_pessoal_grau_relacionamento(Index).SelLength = Len(txt_pessoal_grau_relacionamento(Index))
End Sub
Private Sub txt_pessoal_grau_relacionamento_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        tab_dados.Tab = 2
        txt_salario_base.SetFocus
    End If
End Sub
Private Sub txt_pessoal_nome_GotFocus(Index As Integer)
    txt_pessoal_nome(Index).SelLength = Len(txt_pessoal_nome(Index))
End Sub
Private Sub txt_pessoal_nome_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{Tab}"
    End If
End Sub
Private Sub txt_pessoal_telefone_GotFocus(Index As Integer)
'    txt_pessoal_telefone(Index).SelLength = Len(txt_pessoal_telefone(Index))
End Sub
Private Sub txt_pessoal_telefone_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{Tab}"
    End If
End Sub
Private Sub txt_pis_GotFocus()
    txt_pis.SelLength = Len(txt_pis)
End Sub
Private Sub txt_pis_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{Tab}"
    End If
End Sub
Private Sub txt_reservista_GotFocus()
    txt_reservista.SelLength = Len(txt_reservista)
End Sub
Private Sub txt_reservista_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{Tab}"
    End If
End Sub
Private Sub txt_salario_base_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        chk_adicional_periculosidade.SetFocus
    End If
End Sub
Private Sub txt_salario_base_LostFocus()
    txt_salario_base = Format(txt_salario_base, "###,##0.00")
End Sub
Private Sub txt_senha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub txt_serie_carteira_trabalho_GotFocus()
    txt_serie_carteira_trabalho.SelLength = Len(txt_serie_carteira_trabalho)
End Sub
Private Sub txt_serie_carteira_trabalho_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{Tab}"
    End If
End Sub
Private Sub txt_sexo_GotFocus()
    txt_sexo.SelLength = Len(txt_sexo)
End Sub
Private Sub txt_sexo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{Tab}"
    End If
End Sub
Private Sub txt_sexo_LostFocus()
    txt_sexo = UCase(txt_sexo)
End Sub
Private Sub txt_situacao_GotFocus()
    txt_situacao.SelLength = Len(txt_situacao)
End Sub
Private Sub txt_situacao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{Tab}"
    End If
End Sub
Private Sub txt_situacao_LostFocus()
    txt_situacao = UCase(txt_situacao)
End Sub
Private Sub txt_telefone_GotFocus()
'    txt_telefone.SelLength = Len(txt_telefone)
End Sub
Private Sub txt_telefone_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{Tab}"
    End If
End Sub
Private Sub txt_titulo_eleitor_GotFocus()
    txt_titulo_eleitor.SelLength = Len(txt_titulo_eleitor)
End Sub
Private Sub txt_titulo_eleitor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{Tab}"
    End If
End Sub
Private Sub txt_uf_GotFocus()
    txt_uf.SelLength = Len(txt_uf)
End Sub
Private Sub txt_uf_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{Tab}"
    End If
End Sub
