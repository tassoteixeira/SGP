VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Movimento_Nfce_Auto 
   Caption         =   "Nota Fiscal ao Consumidor Eletrônica"
   ClientHeight    =   9990
   ClientLeft      =   165
   ClientTop       =   -1305
   ClientWidth     =   13785
   Icon            =   "Movimento_Nfce_Auto.frx":0000
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   Picture         =   "Movimento_Nfce_Auto.frx":27A2
   ScaleHeight     =   9990
   ScaleWidth      =   13785
   Begin VB.CommandButton btnAbortaNFCe 
      Caption         =   "Aborta NFCe (F9)"
      Height          =   495
      Left            =   12120
      TabIndex        =   158
      Top             =   9000
      Width           =   1575
   End
   Begin VB.Timer TimerAguarde 
      Enabled         =   0   'False
      Left            =   3360
      Top             =   9600
   End
   Begin VB.Frame frameAguarde 
      Caption         =   "Aguarde..."
      Height          =   4695
      Left            =   8640
      TabIndex        =   154
      Top             =   840
      Visible         =   0   'False
      Width           =   4155
      Begin VB.Label lblTituloAguarde 
         Alignment       =   2  'Center
         Caption         =   "lblTituloAguarde"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   120
         TabIndex        =   157
         Top             =   1080
         Width           =   3975
      End
      Begin VB.Label lblMensagemAguarde 
         Alignment       =   2  'Center
         Caption         =   "lblMensagemAguarde"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   240
         TabIndex        =   156
         Top             =   3000
         Width           =   4035
      End
      Begin VB.Label lblContadorAguarde 
         Alignment       =   2  'Center
         Caption         =   "lblContadorAguarde"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   615
         Left            =   840
         TabIndex        =   155
         Top             =   1920
         Visible         =   0   'False
         Width           =   1575
      End
   End
   Begin VB.Timer TimerIdentFid 
      Enabled         =   0   'False
      Left            =   2580
      Top             =   9540
   End
   Begin MSCommLib.MSComm MSCommIdentFid 
      Left            =   1860
      Top             =   9420
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton cmdIniciaProcessoFinalizacaoNFCe 
      Caption         =   "&Finaliza NFCe (F10)"
      Height          =   495
      Left            =   7920
      TabIndex        =   83
      ToolTipText     =   "Inicia processo de fechamento da NFCe"
      Top             =   9000
      Width           =   2175
   End
   Begin VB.Timer TimerAutomacao 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1200
      Top             =   9480
   End
   Begin VB.Timer Timer2 
      Interval        =   30
      Left            =   720
      Top             =   9480
   End
   Begin RichTextLib.RichTextBox txt_cupom_fiscal 
      Height          =   8835
      Left            =   7920
      TabIndex        =   84
      Top             =   60
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   15584
      _Version        =   393217
      BackColor       =   16777215
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"Movimento_Nfce_Auto.frx":2BE8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame frmDados 
      Height          =   9555
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   7850
      Begin VB.CommandButton cmd_bico 
         Caption         =   "&Bico 48"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   47
         Left            =   1800
         Picture         =   "Movimento_Nfce_Auto.frx":2C68
         Style           =   1  'Graphical
         TabIndex        =   190
         ToolTipText     =   "Bico Livre para Abastecimento."
         Top             =   6480
         Width           =   795
      End
      Begin VB.CommandButton cmd_bico 
         Caption         =   "&Bico 47"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   46
         Left            =   960
         Picture         =   "Movimento_Nfce_Auto.frx":516A
         Style           =   1  'Graphical
         TabIndex        =   188
         ToolTipText     =   "Bico Livre para Abastecimento."
         Top             =   6480
         Width           =   795
      End
      Begin VB.CommandButton cmd_bico 
         Caption         =   "&Bico 46"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   45
         Left            =   90
         Picture         =   "Movimento_Nfce_Auto.frx":766C
         Style           =   1  'Graphical
         TabIndex        =   186
         ToolTipText     =   "Bico Livre para Abastecimento."
         Top             =   6480
         Width           =   795
      End
      Begin VB.CommandButton cmd_bico 
         Caption         =   "&Bico 45"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   44
         Left            =   6960
         Picture         =   "Movimento_Nfce_Auto.frx":9B6E
         Style           =   1  'Graphical
         TabIndex        =   184
         ToolTipText     =   "Bico Livre para Abastecimento."
         Top             =   5400
         Width           =   795
      End
      Begin VB.CommandButton cmd_bico 
         Caption         =   "&Bico 44"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   43
         Left            =   6100
         Picture         =   "Movimento_Nfce_Auto.frx":C070
         Style           =   1  'Graphical
         TabIndex        =   182
         ToolTipText     =   "Bico Livre para Abastecimento."
         Top             =   5400
         Width           =   795
      End
      Begin VB.CommandButton cmd_bico 
         Caption         =   "&Bico 43"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   42
         Left            =   5250
         Picture         =   "Movimento_Nfce_Auto.frx":E572
         Style           =   1  'Graphical
         TabIndex        =   180
         ToolTipText     =   "Bico Livre para Abastecimento."
         Top             =   5400
         Width           =   795
      End
      Begin VB.CommandButton cmd_bico 
         Caption         =   "&Bico 42"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   41
         Left            =   4400
         Picture         =   "Movimento_Nfce_Auto.frx":10A74
         Style           =   1  'Graphical
         TabIndex        =   178
         ToolTipText     =   "Bico Livre para Abastecimento."
         Top             =   5400
         Width           =   795
      End
      Begin VB.CommandButton cmd_bico 
         Caption         =   "&Bico 41"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   40
         Left            =   3520
         Picture         =   "Movimento_Nfce_Auto.frx":12F76
         Style           =   1  'Graphical
         TabIndex        =   176
         ToolTipText     =   "Bico Livre para Abastecimento."
         Top             =   5400
         Width           =   795
      End
      Begin VB.CommandButton cmd_bico 
         Caption         =   "&Bico 40"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   39
         Left            =   2670
         Picture         =   "Movimento_Nfce_Auto.frx":15478
         Style           =   1  'Graphical
         TabIndex        =   174
         ToolTipText     =   "Bico Livre para Abastecimento."
         Top             =   5400
         Width           =   795
      End
      Begin VB.CommandButton cmd_bico 
         Caption         =   "&Bico 39"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   38
         Left            =   1820
         Picture         =   "Movimento_Nfce_Auto.frx":1797A
         Style           =   1  'Graphical
         TabIndex        =   172
         ToolTipText     =   "Bico Livre para Abastecimento."
         Top             =   5400
         Width           =   795
      End
      Begin VB.CommandButton cmd_bico 
         Caption         =   "&Bico 38"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   37
         Left            =   940
         Picture         =   "Movimento_Nfce_Auto.frx":19E7C
         Style           =   1  'Graphical
         TabIndex        =   170
         ToolTipText     =   "Bico Livre para Abastecimento."
         Top             =   5400
         Width           =   795
      End
      Begin VB.CommandButton cmd_bico 
         Caption         =   "&Bico 37"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   36
         Left            =   90
         Picture         =   "Movimento_Nfce_Auto.frx":1C37E
         Style           =   1  'Graphical
         TabIndex        =   168
         ToolTipText     =   "Bico Livre para Abastecimento."
         Top             =   5400
         Width           =   795
      End
      Begin VB.CommandButton cmd_bico 
         Caption         =   "&Bico 36"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   35
         Left            =   6960
         Picture         =   "Movimento_Nfce_Auto.frx":1E880
         Style           =   1  'Graphical
         TabIndex        =   166
         ToolTipText     =   "Bico Livre para Abastecimento."
         Top             =   4250
         Width           =   795
      End
      Begin VB.CommandButton cmd_bico 
         Caption         =   "&Bico 35"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   34
         Left            =   6100
         Picture         =   "Movimento_Nfce_Auto.frx":20D82
         Style           =   1  'Graphical
         TabIndex        =   164
         ToolTipText     =   "Bico Livre para Abastecimento."
         Top             =   4250
         Width           =   795
      End
      Begin VB.CommandButton cmd_bico 
         Caption         =   "&Bico 34"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   33
         Left            =   5250
         Picture         =   "Movimento_Nfce_Auto.frx":23284
         Style           =   1  'Graphical
         TabIndex        =   162
         ToolTipText     =   "Bico Livre para Abastecimento."
         Top             =   4250
         Width           =   795
      End
      Begin VB.CommandButton cmd_bico 
         Caption         =   "&Bico 33"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   32
         Left            =   4400
         Picture         =   "Movimento_Nfce_Auto.frx":25786
         Style           =   1  'Graphical
         TabIndex        =   160
         ToolTipText     =   "Bico Livre para Abastecimento."
         Top             =   4250
         Width           =   795
      End
      Begin VB.CommandButton cmdAtualizaAbastecimento 
         Caption         =   "A&tualizar Abastecimentos"
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
         Left            =   240
         MaskColor       =   &H000000FF&
         TabIndex        =   153
         ToolTipText     =   "Visualiza os Abastecimentos Não Recebidos."
         Top             =   7560
         UseMaskColor    =   -1  'True
         Width           =   2295
      End
      Begin VB.CommandButton cmd_bico 
         Caption         =   "&Bico 32"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   31
         Left            =   3520
         Picture         =   "Movimento_Nfce_Auto.frx":27C88
         Style           =   1  'Graphical
         TabIndex        =   67
         ToolTipText     =   "Bico Livre para Abastecimento."
         Top             =   4250
         Width           =   795
      End
      Begin VB.CommandButton cmd_bico 
         Caption         =   "&Bico 31"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   30
         Left            =   2670
         Picture         =   "Movimento_Nfce_Auto.frx":2A18A
         Style           =   1  'Graphical
         TabIndex        =   65
         ToolTipText     =   "Bico Livre para Abastecimento."
         Top             =   4250
         Width           =   795
      End
      Begin VB.CommandButton cmd_bico 
         Caption         =   "&Bico 30"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   29
         Left            =   1820
         Picture         =   "Movimento_Nfce_Auto.frx":2C68C
         Style           =   1  'Graphical
         TabIndex        =   63
         ToolTipText     =   "Bico Livre para Abastecimento."
         Top             =   4250
         Width           =   795
      End
      Begin VB.CommandButton cmd_bico 
         Caption         =   "&Bico 29"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   28
         Left            =   940
         Picture         =   "Movimento_Nfce_Auto.frx":2EB8E
         Style           =   1  'Graphical
         TabIndex        =   61
         ToolTipText     =   "Bico Livre para Abastecimento."
         Top             =   4250
         Width           =   795
      End
      Begin VB.CommandButton cmd_bico 
         Caption         =   "&Bico 28"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   27
         Left            =   90
         Picture         =   "Movimento_Nfce_Auto.frx":31090
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Bico Livre para Abastecimento."
         Top             =   4250
         Width           =   795
      End
      Begin VB.CommandButton cmd_bico 
         Caption         =   "&Bico 27"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   26
         Left            =   6960
         Picture         =   "Movimento_Nfce_Auto.frx":33592
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Bico Livre para Abastecimento."
         Top             =   3120
         Width           =   795
      End
      Begin VB.CommandButton cmd_bico 
         Caption         =   "&Bico 26"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   25
         Left            =   6100
         Picture         =   "Movimento_Nfce_Auto.frx":35A94
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Bico Livre para Abastecimento."
         Top             =   3120
         Width           =   795
      End
      Begin VB.CommandButton cmd_bico 
         Caption         =   "&Bico 25"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   24
         Left            =   5250
         Picture         =   "Movimento_Nfce_Auto.frx":37F96
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Bico Livre para Abastecimento."
         Top             =   3120
         Width           =   795
      End
      Begin VB.CommandButton cmdPesquisa 
         Caption         =   "&Pesquisa"
         Height          =   255
         Left            =   4500
         TabIndex        =   76
         Top             =   8280
         Width           =   1035
      End
      Begin VB.CommandButton cmd_bico 
         Caption         =   "&Bico 24"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   23
         Left            =   4400
         Picture         =   "Movimento_Nfce_Auto.frx":3A498
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Bico Livre para Abastecimento."
         Top             =   3120
         Width           =   795
      End
      Begin VB.CommandButton cmd_bico 
         Caption         =   "&Bico 23"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   22
         Left            =   3520
         Picture         =   "Movimento_Nfce_Auto.frx":3C99A
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Bico Livre para Abastecimento."
         Top             =   3120
         Width           =   795
      End
      Begin VB.CommandButton cmd_bico 
         Caption         =   "&Bico 22"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   21
         Left            =   2670
         Picture         =   "Movimento_Nfce_Auto.frx":3EE9C
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Bico Livre para Abastecimento."
         Top             =   3120
         Width           =   795
      End
      Begin VB.CommandButton cmd_bico 
         Caption         =   "&Bico 21"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   20
         Left            =   1820
         Picture         =   "Movimento_Nfce_Auto.frx":4139E
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Bico Livre para Abastecimento."
         Top             =   3120
         Width           =   795
      End
      Begin VB.CommandButton cmd_bico 
         Caption         =   "&Bico 20"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   19
         Left            =   940
         Picture         =   "Movimento_Nfce_Auto.frx":438A0
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Bico Livre para Abastecimento."
         Top             =   3120
         Width           =   795
      End
      Begin VB.CommandButton cmd_bico 
         Caption         =   "&Bico 19"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   18
         Left            =   90
         Picture         =   "Movimento_Nfce_Auto.frx":45DA2
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Bico Livre para Abastecimento."
         Top             =   3120
         Width           =   795
      End
      Begin VB.CommandButton cmd_bico 
         Caption         =   "&Bico 18"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   17
         Left            =   6960
         Picture         =   "Movimento_Nfce_Auto.frx":482A4
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Bico Livre para Abastecimento."
         Top             =   1980
         Width           =   795
      End
      Begin VB.CommandButton cmd_bico 
         Caption         =   "&Bico 17"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   16
         Left            =   6100
         Picture         =   "Movimento_Nfce_Auto.frx":4A7A6
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Bico Livre para Abastecimento."
         Top             =   1980
         Width           =   795
      End
      Begin VB.CommandButton cmd_bico 
         Caption         =   "&Bico 16"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   15
         Left            =   5250
         Picture         =   "Movimento_Nfce_Auto.frx":4CCA8
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Bico Livre para Abastecimento."
         Top             =   1980
         Width           =   795
      End
      Begin VB.CommandButton cmd_bico 
         Caption         =   "&Bico 15"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   14
         Left            =   4400
         Picture         =   "Movimento_Nfce_Auto.frx":4F1AA
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Bico Livre para Abastecimento."
         Top             =   1980
         Width           =   795
      End
      Begin VB.CommandButton cmd_bico 
         Caption         =   "&Bico 14"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   13
         Left            =   3520
         Picture         =   "Movimento_Nfce_Auto.frx":516AC
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Bico Livre para Abastecimento."
         Top             =   1980
         Width           =   795
      End
      Begin VB.CommandButton cmd_bico 
         Caption         =   "&Bico 13"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   12
         Left            =   2670
         Picture         =   "Movimento_Nfce_Auto.frx":53BAE
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Bico Livre para Abastecimento."
         Top             =   1980
         Width           =   795
      End
      Begin VB.CommandButton cmd_bico 
         Caption         =   "&Bico 12"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   11
         Left            =   1820
         Picture         =   "Movimento_Nfce_Auto.frx":560B0
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Bico Livre para Abastecimento."
         Top             =   1980
         Width           =   795
      End
      Begin VB.CommandButton cmd_bico 
         Caption         =   "&Bico 11"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   10
         Left            =   940
         Picture         =   "Movimento_Nfce_Auto.frx":585B2
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Bico Livre para Abastecimento."
         Top             =   1980
         Width           =   795
      End
      Begin VB.CommandButton cmd_bico 
         Caption         =   "&Bico 10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   9
         Left            =   90
         Picture         =   "Movimento_Nfce_Auto.frx":5AAB4
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Bico Livre para Abastecimento."
         Top             =   1980
         Width           =   795
      End
      Begin VB.CommandButton cmd_bico 
         Caption         =   "&Bico 09"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   8
         Left            =   6960
         Picture         =   "Movimento_Nfce_Auto.frx":5CFB6
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Bico Livre para Abastecimento."
         Top             =   840
         Width           =   795
      End
      Begin VB.ComboBox cboTipoSubEstoque 
         Height          =   315
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   71
         Top             =   7980
         Width           =   2175
      End
      Begin VB.CommandButton cmd_bico 
         Caption         =   "&Bico 08"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   7
         Left            =   6100
         Picture         =   "Movimento_Nfce_Auto.frx":5F4B8
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Bico Livre para Abastecimento."
         Top             =   840
         Width           =   795
      End
      Begin VB.CommandButton cmd_bico 
         Caption         =   "&Bico 07"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   6
         Left            =   5250
         Picture         =   "Movimento_Nfce_Auto.frx":619BA
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Bico Livre para Abastecimento."
         Top             =   840
         Width           =   795
      End
      Begin VB.CommandButton cmd_bico 
         Caption         =   "&Bico 06"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   5
         Left            =   4400
         Picture         =   "Movimento_Nfce_Auto.frx":63EBC
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Bico Livre para Abastecimento."
         Top             =   840
         Width           =   795
      End
      Begin VB.CommandButton cmd_bico 
         Caption         =   "&Bico 04"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   3
         Left            =   2670
         Picture         =   "Movimento_Nfce_Auto.frx":663BE
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Bico Livre para Abastecimento."
         Top             =   840
         Width           =   795
      End
      Begin VB.CommandButton cmd_bico 
         Caption         =   "&Bico 05"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   4
         Left            =   3520
         Picture         =   "Movimento_Nfce_Auto.frx":688C0
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Bico Livre para Abastecimento."
         Top             =   840
         Width           =   795
      End
      Begin VB.CommandButton cmd_abastecimentos_nao_recebidos 
         Caption         =   "&Abastecimentos Não Recebidos"
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
         Left            =   2640
         MaskColor       =   &H000000FF&
         TabIndex        =   69
         ToolTipText     =   "Visualiza os Abastecimentos Não Recebidos."
         Top             =   7560
         UseMaskColor    =   -1  'True
         Width           =   3015
      End
      Begin VB.CommandButton cmd_bico 
         Caption         =   "&Bico 03"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   2
         Left            =   1820
         Picture         =   "Movimento_Nfce_Auto.frx":6ADC2
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Bico Livre para Abastecimento."
         Top             =   840
         Width           =   795
      End
      Begin VB.CommandButton cmd_bico 
         Caption         =   "&Bico 02"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   1
         Left            =   940
         Picture         =   "Movimento_Nfce_Auto.frx":6D2C4
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Bico Livre para Abastecimento."
         Top             =   840
         Width           =   795
      End
      Begin VB.CommandButton cmd_bico 
         Caption         =   "&Bico 01"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   0
         Left            =   90
         Picture         =   "Movimento_Nfce_Auto.frx":6F7C6
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Bico Livre para Abastecimento."
         Top             =   840
         Width           =   795
      End
      Begin VB.TextBox txt_quantidade 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   80
         Top             =   9180
         Width           =   1095
      End
      Begin VB.TextBox txt_valor_unitario 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   78
         Top             =   9180
         Width           =   1095
      End
      Begin VB.TextBox txt_produto 
         Height          =   300
         Left            =   120
         MaxLength       =   18
         TabIndex        =   73
         Top             =   8580
         Width           =   795
      End
      Begin VB.TextBox txt_cliente 
         Height          =   285
         Left            =   120
         MaxLength       =   6
         TabIndex        =   2
         Top             =   360
         Width           =   795
      End
      Begin VB.TextBox txt_valor_total 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4560
         MaxLength       =   10
         TabIndex        =   82
         Top             =   9180
         Width           =   1095
      End
      Begin MSAdodcLib.Adodc adodcCliente 
         Height          =   330
         Left            =   2520
         Top             =   360
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
         Caption         =   "adodcCliente"
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
      Begin MSDataListLib.DataCombo dtcboCliente 
         Bindings        =   "Movimento_Nfce_Auto.frx":71CC8
         Height          =   315
         Left            =   1020
         TabIndex        =   4
         Top             =   360
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Razao Social"
         BoundColumn     =   "Codigo"
         Text            =   "dtcboCliente"
      End
      Begin MSAdodcLib.Adodc adodcProduto 
         Height          =   330
         Left            =   2280
         Top             =   8580
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
         Caption         =   "adodcProduto"
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
      Begin MSDataListLib.DataCombo dtcboProduto 
         Bindings        =   "Movimento_Nfce_Auto.frx":71CE3
         Height          =   315
         Left            =   960
         TabIndex        =   75
         Top             =   8580
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   "dtcboProduto"
      End
      Begin VB.Label lbl_automacao_valor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   47
         Left            =   1800
         TabIndex        =   191
         Top             =   7200
         Width           =   795
      End
      Begin VB.Label lbl_automacao_valor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   46
         Left            =   960
         TabIndex        =   189
         Top             =   7200
         Width           =   795
      End
      Begin VB.Label lbl_automacao_valor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   45
         Left            =   90
         TabIndex        =   187
         Top             =   7200
         Width           =   795
      End
      Begin VB.Label lbl_automacao_valor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   44
         Left            =   6960
         TabIndex        =   185
         Top             =   6120
         Width           =   795
      End
      Begin VB.Label lbl_automacao_valor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   43
         Left            =   6100
         TabIndex        =   183
         Top             =   6120
         Width           =   795
      End
      Begin VB.Label lbl_automacao_valor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   42
         Left            =   5250
         TabIndex        =   181
         Top             =   6120
         Width           =   795
      End
      Begin VB.Label lbl_automacao_valor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   41
         Left            =   4400
         TabIndex        =   179
         Top             =   6120
         Width           =   795
      End
      Begin VB.Label lbl_automacao_valor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   40
         Left            =   3520
         TabIndex        =   177
         Top             =   6120
         Width           =   795
      End
      Begin VB.Label lbl_automacao_valor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   39
         Left            =   2670
         TabIndex        =   175
         Top             =   6120
         Width           =   795
      End
      Begin VB.Label lbl_automacao_valor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   38
         Left            =   1820
         TabIndex        =   173
         Top             =   6120
         Width           =   795
      End
      Begin VB.Label lbl_automacao_valor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   37
         Left            =   940
         TabIndex        =   171
         Top             =   6120
         Width           =   795
      End
      Begin VB.Label lbl_automacao_valor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   36
         Left            =   90
         TabIndex        =   169
         Top             =   6120
         Width           =   795
      End
      Begin VB.Label lbl_automacao_valor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   35
         Left            =   6960
         TabIndex        =   167
         Top             =   4965
         Width           =   795
      End
      Begin VB.Label lbl_automacao_valor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   34
         Left            =   6100
         TabIndex        =   165
         Top             =   4965
         Width           =   795
      End
      Begin VB.Label lbl_automacao_valor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   33
         Left            =   5250
         TabIndex        =   163
         Top             =   4965
         Width           =   795
      End
      Begin VB.Label lbl_automacao_valor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   32
         Left            =   4400
         TabIndex        =   161
         Top             =   4965
         Width           =   795
      End
      Begin VB.Label lbl_automacao_valor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   31
         Left            =   3520
         TabIndex        =   68
         Top             =   4965
         Width           =   795
      End
      Begin VB.Label lbl_automacao_valor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   30
         Left            =   2670
         TabIndex        =   66
         Top             =   4965
         Width           =   795
      End
      Begin VB.Label lbl_automacao_valor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   29
         Left            =   1820
         TabIndex        =   64
         Top             =   4965
         Width           =   795
      End
      Begin VB.Label lbl_automacao_valor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   28
         Left            =   940
         TabIndex        =   62
         Top             =   4965
         Width           =   795
      End
      Begin VB.Label lbl_automacao_valor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   27
         Left            =   90
         TabIndex        =   60
         Top             =   4965
         Width           =   795
      End
      Begin VB.Label lbl_automacao_valor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   26
         Left            =   6960
         TabIndex        =   58
         Top             =   3840
         Width           =   795
      End
      Begin VB.Label lbl_automacao_valor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   25
         Left            =   6100
         TabIndex        =   56
         Top             =   3840
         Width           =   795
      End
      Begin VB.Label lbl_automacao_valor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   24
         Left            =   5250
         TabIndex        =   54
         Top             =   3840
         Width           =   795
      End
      Begin VB.Label lbl_automacao_valor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   23
         Left            =   4400
         TabIndex        =   52
         Top             =   3840
         Width           =   795
      End
      Begin VB.Label lbl_automacao_valor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   22
         Left            =   3520
         TabIndex        =   50
         Top             =   3840
         Width           =   795
      End
      Begin VB.Label lbl_automacao_valor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   21
         Left            =   2670
         TabIndex        =   48
         Top             =   3840
         Width           =   795
      End
      Begin VB.Label lbl_automacao_valor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   20
         Left            =   1820
         TabIndex        =   46
         Top             =   3840
         Width           =   795
      End
      Begin VB.Label lbl_automacao_valor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   19
         Left            =   940
         TabIndex        =   44
         Top             =   3840
         Width           =   795
      End
      Begin VB.Label lbl_automacao_valor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   18
         Left            =   90
         TabIndex        =   42
         Top             =   3840
         Width           =   795
      End
      Begin VB.Label lbl_automacao_valor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   17
         Left            =   6960
         TabIndex        =   40
         Top             =   2700
         Width           =   795
      End
      Begin VB.Label lbl_automacao_valor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   16
         Left            =   6100
         TabIndex        =   38
         Top             =   2700
         Width           =   795
      End
      Begin VB.Label lbl_automacao_valor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   15
         Left            =   5250
         TabIndex        =   36
         Top             =   2700
         Width           =   795
      End
      Begin VB.Label lbl_automacao_valor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   14
         Left            =   4400
         TabIndex        =   34
         Top             =   2700
         Width           =   795
      End
      Begin VB.Label lbl_automacao_valor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   13
         Left            =   3520
         TabIndex        =   32
         Top             =   2700
         Width           =   795
      End
      Begin VB.Label lbl_automacao_valor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   12
         Left            =   2670
         TabIndex        =   30
         Top             =   2700
         Width           =   795
      End
      Begin VB.Label lbl_automacao_valor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   11
         Left            =   1820
         TabIndex        =   28
         Top             =   2700
         Width           =   795
      End
      Begin VB.Label lbl_automacao_valor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   10
         Left            =   940
         TabIndex        =   26
         Top             =   2700
         Width           =   795
      End
      Begin VB.Label lbl_automacao_valor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   9
         Left            =   90
         TabIndex        =   24
         Top             =   2700
         Width           =   795
      End
      Begin VB.Label lbl_automacao_valor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   8
         Left            =   6960
         TabIndex        =   22
         Top             =   1560
         Width           =   795
      End
      Begin VB.Label lblTipoSubEstoque 
         Caption         =   "&Tipo do Sub-Estoque"
         Height          =   315
         Left            =   1680
         TabIndex        =   70
         Top             =   7980
         Width           =   1755
      End
      Begin VB.Label lbl_automacao_valor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   7
         Left            =   6100
         TabIndex        =   20
         Top             =   1560
         Width           =   795
      End
      Begin VB.Label lbl_automacao_valor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   6
         Left            =   5250
         TabIndex        =   18
         Top             =   1560
         Width           =   795
      End
      Begin VB.Label lbl_automacao_valor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   5
         Left            =   4400
         TabIndex        =   16
         Top             =   1560
         Width           =   795
      End
      Begin VB.Label lbl_automacao_valor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   4
         Left            =   3520
         TabIndex        =   14
         Top             =   1560
         Width           =   795
      End
      Begin VB.Label lbl_automacao_valor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   3
         Left            =   2670
         TabIndex        =   12
         Top             =   1560
         Width           =   795
      End
      Begin VB.Label lbl_automacao_valor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   2
         Left            =   1820
         TabIndex        =   10
         Top             =   1560
         Width           =   795
      End
      Begin VB.Label lbl_automacao_valor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   1
         Left            =   940
         TabIndex        =   8
         Top             =   1560
         Width           =   795
      End
      Begin VB.Label lbl_automacao_valor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   0
         Left            =   90
         TabIndex        =   6
         Top             =   1560
         Width           =   795
      End
      Begin VB.Label Label3 
         Caption         =   "Có&digo"
         Height          =   315
         Index           =   17
         Left            =   120
         TabIndex        =   72
         Top             =   8340
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "No&me do Cliente"
         Height          =   315
         Index           =   13
         Left            =   1020
         TabIndex        =   3
         Top             =   120
         Width           =   1395
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   5750
         Y1              =   7905
         Y2              =   7905
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   5750
         Y1              =   750
         Y2              =   750
      End
      Begin VB.Label Label3 
         Caption         =   "&Quantidade"
         Height          =   315
         Index           =   2
         Left            =   2400
         TabIndex        =   79
         Top             =   8940
         Width           =   915
      End
      Begin VB.Label Label3 
         Caption         =   "Preço &unitário"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   77
         Top             =   8940
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Nome do P&roduto"
         Height          =   315
         Index           =   0
         Left            =   960
         TabIndex        =   74
         Top             =   8340
         Width           =   1395
      End
      Begin VB.Label Label3 
         Caption         =   "&Código"
         Height          =   315
         Index           =   9
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Pr&eço total"
         Height          =   315
         Index           =   5
         Left            =   4560
         TabIndex        =   81
         Top             =   8940
         Width           =   855
      End
   End
   Begin VB.Frame frm_fechamento_cupom 
      Caption         =   "Fechamento da NFCe"
      Height          =   4395
      Left            =   120
      TabIndex        =   95
      Top             =   1920
      Width           =   7460
      Begin VB.CommandButton btnDescontoCartaoEspecial 
         Caption         =   "Desconto Cartão Cliente Especial"
         Height          =   615
         Left            =   5760
         TabIndex        =   159
         Top             =   1800
         Width           =   1575
      End
      Begin VB.CommandButton cmdCartaoFidelidadeDesconto 
         Caption         =   "Desconto Cartão Fidelidade"
         Height          =   615
         Left            =   5760
         TabIndex        =   152
         ToolTipText     =   "Calcula Desconto com Cartão de Desconto/Fidelidade"
         Top             =   3240
         Width           =   1575
      End
      Begin VB.CommandButton cmd_DescontoCorreio 
         Caption         =   "Desconto Cartão dos &Correios"
         Height          =   615
         Left            =   5760
         TabIndex        =   151
         ToolTipText     =   "Calcula Desconto com Cartão dos Correios"
         Top             =   2520
         Width           =   1575
      End
      Begin VB.CommandButton cmdDescontoAppPostoAki 
         Caption         =   "Desconto App PostoAki"
         Height          =   615
         Left            =   4080
         TabIndex        =   150
         ToolTipText     =   "Calcula Desconto com App PostoAki"
         Top             =   3240
         Width           =   1575
      End
      Begin VB.CommandButton cmdInformaPlacaVeiculo 
         Caption         =   "&Informa Placa e KM do Veículo"
         Height          =   315
         Left            =   3180
         TabIndex        =   102
         Top             =   1500
         Width           =   2535
      End
      Begin VB.TextBox txt_observacao 
         Height          =   285
         Left            =   120
         MaxLength       =   48
         TabIndex        =   104
         Top             =   1800
         Width           =   5595
      End
      Begin VB.TextBox txt_observacao_2 
         Height          =   285
         Left            =   120
         MaxLength       =   48
         TabIndex        =   105
         Top             =   2100
         Width           =   5595
      End
      Begin VB.TextBox txt_numero_nota_abastecimento 
         Height          =   285
         Left            =   3180
         MaxLength       =   6
         TabIndex        =   139
         Top             =   3300
         Width           =   795
      End
      Begin VB.TextBox txt_kilometragem 
         Height          =   285
         Left            =   1680
         MaxLength       =   12
         TabIndex        =   117
         Top             =   3300
         Width           =   1215
      End
      Begin VB.TextBox txt_placa 
         Height          =   285
         Left            =   120
         MaxLength       =   8
         TabIndex        =   115
         Top             =   3300
         Width           =   975
      End
      Begin VB.ComboBox cbo_forma_pagamento 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   97
         Top             =   480
         Width           =   3195
      End
      Begin VB.CommandButton cmdIniciaProcessoNFCe 
         Caption         =   "O&K"
         Height          =   375
         Left            =   4860
         TabIndex        =   123
         ToolTipText     =   "Confirma o fechamento deste cupom"
         Top             =   3900
         Width           =   855
      End
      Begin VB.CommandButton cmd_cancelar2 
         Caption         =   "C&ancelar"
         Height          =   375
         Left            =   3900
         TabIndex        =   122
         ToolTipText     =   "Cancela o fechamento deste cupom"
         Top             =   3900
         Width           =   855
      End
      Begin VB.TextBox txt_numero_cheque 
         Height          =   285
         Left            =   120
         MaxLength       =   6
         TabIndex        =   119
         Top             =   3960
         Width           =   795
      End
      Begin VB.TextBox txt_telefone 
         Height          =   285
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   121
         Top             =   3960
         Width           =   1455
      End
      Begin VB.TextBox txt_valor_recebido 
         Height          =   285
         Left            =   3180
         MaxLength       =   10
         TabIndex        =   111
         Top             =   2700
         Width           =   1095
      End
      Begin VB.TextBox txt_valor_desconto 
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   107
         Top             =   2700
         Width           =   1095
      End
      Begin VB.TextBox txt_nome_cliente 
         Height          =   285
         Left            =   120
         MaxLength       =   40
         TabIndex        =   101
         Top             =   1140
         Width           =   5595
      End
      Begin VB.TextBox txt_cpf 
         Height          =   285
         Left            =   3420
         MaxLength       =   20
         TabIndex        =   99
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "&Observações:"
         Height          =   195
         Index           =   20
         Left            =   120
         TabIndex        =   103
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label13 
         Caption         =   "Número da Nota de Abastecimento"
         Height          =   195
         Left            =   3180
         TabIndex        =   140
         Top             =   3060
         Width           =   2535
      End
      Begin VB.Label Label12 
         Caption         =   "K&ilometragem"
         Height          =   195
         Left            =   1680
         TabIndex        =   116
         Top             =   3060
         Width           =   1395
      End
      Begin VB.Label Label11 
         Caption         =   "&Placa do Veículo"
         Height          =   195
         Left            =   120
         TabIndex        =   114
         Top             =   3060
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "&Forma de Pagamento"
         Height          =   195
         Index           =   12
         Left            =   120
         TabIndex        =   96
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lbl_numero_cheque 
         Caption         =   "Número do Cheque"
         Height          =   195
         Left            =   120
         TabIndex        =   118
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Label lbl_telefone 
         Caption         =   "Telefone"
         Height          =   195
         Left            =   1800
         TabIndex        =   120
         Top             =   3720
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Valor da Compra"
         Height          =   195
         Left            =   1680
         TabIndex        =   108
         Top             =   2460
         Width           =   1335
      End
      Begin VB.Label lbl_valor_compra 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1680
         TabIndex        =   109
         Top             =   2700
         Width           =   1095
      End
      Begin VB.Label lbl_valor_troco1 
         Caption         =   "Valor do Troco"
         Height          =   195
         Left            =   4620
         TabIndex        =   112
         Top             =   2460
         Width           =   1095
      End
      Begin VB.Label lbl_valor_troco 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4620
         TabIndex        =   113
         Top             =   2700
         Width           =   1095
      End
      Begin VB.Label lbl_valor_recebido 
         Caption         =   "Valor &Recebido"
         Height          =   195
         Left            =   3180
         TabIndex        =   110
         Top             =   2460
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Valor do &Desconto"
         Height          =   195
         Left            =   120
         TabIndex        =   106
         Top             =   2460
         Width           =   1395
      End
      Begin VB.Label Label3 
         Caption         =   "&Nome do Cliente"
         Height          =   195
         Index           =   18
         Left            =   120
         TabIndex        =   100
         Top             =   900
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "CPF/CNP&J"
         Height          =   195
         Index           =   19
         Left            =   3420
         TabIndex        =   98
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame frmDescarregar 
      Caption         =   "Descarregar Abastecimento"
      Height          =   2595
      Left            =   0
      TabIndex        =   141
      Top             =   660
      Width           =   5055
      Begin VB.CheckBox chkDesconto 
         Height          =   315
         Left            =   2400
         TabIndex        =   147
         Top             =   1560
         Width           =   435
      End
      Begin VB.TextBox txtQuantidadeDescarregamento 
         Height          =   315
         Left            =   2400
         MaxLength       =   2
         TabIndex        =   145
         Top             =   1020
         Width           =   315
      End
      Begin VB.ComboBox cboBico 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   143
         Top             =   540
         Width           =   555
      End
      Begin VB.CommandButton cmdCancelarDescarregar 
         Caption         =   "&Cancelar"
         Height          =   315
         Left            =   3120
         TabIndex        =   149
         ToolTipText     =   "Voltar para visualização dos bicos."
         Top             =   2040
         Width           =   795
      End
      Begin VB.CommandButton cmdOkDescarregar 
         Caption         =   "&Ok"
         Height          =   315
         Left            =   1560
         TabIndex        =   148
         ToolTipText     =   "Voltar para visualização dos bicos."
         Top             =   2040
         Width           =   795
      End
      Begin VB.Label Label16 
         Caption         =   "Conceder Desconto"
         Height          =   315
         Left            =   180
         TabIndex        =   146
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label Label15 
         Caption         =   "Quantidade de Abastecimento"
         Height          =   315
         Left            =   180
         TabIndex        =   144
         Top             =   1020
         Width           =   2175
      End
      Begin VB.Label Label14 
         Caption         =   "Código do Bico"
         Height          =   315
         Left            =   180
         TabIndex        =   142
         Top             =   540
         Width           =   2175
      End
   End
   Begin VB.Frame frm_ponto 
      Caption         =   "Identificação de Funcionário"
      Height          =   9500
      Left            =   30
      TabIndex        =   86
      Top             =   60
      Width           =   7850
      Begin VB.TextBox txt_senha_ponto 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   92
         Top             =   3840
         Width           =   1215
      End
      Begin VB.CommandButton cmd_ok_ponto 
         Caption         =   "&OK"
         Height          =   375
         Left            =   4800
         TabIndex        =   94
         ToolTipText     =   "Confirma este registro de ponto de funcionário."
         Top             =   3780
         Width           =   855
      End
      Begin VB.CommandButton cmd_cancelar_ponto 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   3840
         TabIndex        =   93
         ToolTipText     =   "Cancela este registro de ponto de funcionário."
         Top             =   3780
         Width           =   855
      End
      Begin VB.TextBox txt_funcionario_ponto 
         Height          =   285
         Left            =   120
         MaxLength       =   3
         TabIndex        =   88
         Top             =   2400
         Width           =   555
      End
      Begin MSAdodcLib.Adodc adodcFuncionario 
         Height          =   330
         Left            =   2640
         Top             =   3060
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
      Begin MSDataListLib.DataCombo dtcboFuncionario 
         Bindings        =   "Movimento_Nfce_Auto.frx":71CFE
         Height          =   315
         Left            =   120
         TabIndex        =   90
         Top             =   3060
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   "dtcboFuncionario"
      End
      Begin VB.Label Label3 
         Caption         =   "Có&digo do Funcionário"
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   87
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "&Senha"
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   91
         Top             =   3600
         Width           =   555
      End
      Begin VB.Label Label3 
         Caption         =   "&Nome do Funcionário"
         Height          =   315
         Index           =   14
         Left            =   120
         TabIndex        =   89
         Top             =   2820
         Width           =   1815
      End
   End
   Begin VB.Frame frm_fila_bico 
      Enabled         =   0   'False
      Height          =   8775
      Left            =   30
      TabIndex        =   124
      Top             =   60
      Visible         =   0   'False
      Width           =   7775
      Begin VB.CommandButton cmdDescarregar 
         Caption         =   "&Descarregar"
         Height          =   315
         Left            =   1080
         TabIndex        =   135
         ToolTipText     =   "Descarregar abastecimento de um bico."
         Top             =   4680
         Width           =   1095
      End
      Begin VB.CommandButton cmd_fila_sair 
         Caption         =   "&Sair"
         Height          =   315
         Left            =   3720
         TabIndex        =   137
         ToolTipText     =   "Voltar para visualização dos bicos."
         Top             =   4680
         Width           =   555
      End
      Begin VB.CommandButton cmd_fila_ok 
         Caption         =   "&Imprime Cupom"
         Height          =   315
         Left            =   2280
         TabIndex        =   136
         ToolTipText     =   "Imprime o cupom do ítem selecionado."
         Top             =   4680
         Width           =   1335
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid 
         Height          =   2715
         Left            =   60
         TabIndex        =   133
         Top             =   540
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   4789
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "Abastecimentos Não Recebidos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   60
         TabIndex        =   138
         Top             =   180
         Width           =   5655
      End
      Begin VB.Label lbl_fila_litros 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "88888888"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   3060
         TabIndex        =   134
         Top             =   3735
         Width           =   1035
      End
      Begin VB.Label lbl_fila_total 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "88888888"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   3060
         TabIndex        =   132
         Top             =   4035
         Width           =   1035
      End
      Begin VB.Label Label9 
         BackColor       =   &H00000000&
         Caption         =   "Label9"
         Height          =   255
         Left            =   3180
         TabIndex        =   131
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label Label8 
         BackColor       =   &H00000000&
         Caption         =   "Label8"
         Height          =   255
         Left            =   3180
         TabIndex        =   130
         Top             =   4020
         UseMnemonic     =   0   'False
         Width           =   975
      End
      Begin VB.Label lbl_fila_valor 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "88888888"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   3060
         TabIndex        =   129
         Top             =   3420
         Width           =   1035
      End
      Begin VB.Label Label7 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3180
         TabIndex        =   128
         Top             =   3405
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2460
         TabIndex        =   127
         Top             =   4020
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Valor"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2460
         TabIndex        =   126
         Top             =   3420
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Litros"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2460
         TabIndex        =   125
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Line Line6 
         X1              =   2295
         X2              =   4260
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Line Line5 
         X1              =   4260
         X2              =   4260
         Y1              =   3360
         Y2              =   4335
      End
      Begin VB.Line Line4 
         X1              =   2295
         X2              =   2295
         Y1              =   3360
         Y2              =   4335
      End
      Begin VB.Line Line3 
         X1              =   2295
         X2              =   4260
         Y1              =   3360
         Y2              =   3360
      End
   End
   Begin VB.Label lbl_mensagem 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   50
      TabIndex        =   85
      Top             =   9610
      Width           =   13695
   End
   Begin VB.Menu mnuCaixaPista 
      Caption         =   "Caixa de Pista"
   End
   Begin VB.Menu mnuConsulta 
      Caption         =   "Consultas"
      Begin VB.Menu mnuConsultaCheque 
         Caption         =   "&Cliente de Cheque"
      End
      Begin VB.Menu mnuVisualizaVenda 
         Caption         =   "&Visualiza Vendas"
      End
   End
   Begin VB.Menu mnRelatorio 
      Caption         =   "Relatórios"
      Begin VB.Menu mnAbastecimentoAutomacaoFuncionario 
         Caption         =   "Abastecimento Automação Funcionário"
      End
      Begin VB.Menu mnFaltaCaixa 
         Caption         =   "Falta de Caixa"
      End
   End
   Begin VB.Menu mnuFuncaoADM 
      Caption         =   "Função ADM"
   End
   Begin VB.Menu mnuFuncao 
      Caption         =   "Outras Funções"
      Begin VB.Menu mnuTCS 
         Caption         =   "&Atualização de Preço TicketCar Smart"
      End
      Begin VB.Menu mnuCalculadora 
         Caption         =   "&Calculadora"
      End
      Begin VB.Menu mnuCancelaCartao 
         Caption         =   "&Cancelamento de Cartão"
      End
      Begin VB.Menu mnuFechamentoCaixa 
         Caption         =   "&Fechamento de Caixa"
      End
      Begin VB.Menu mnuLancamentoEncerrante 
         Caption         =   "&Lançamento dos Encerrantes (Automático)"
      End
      Begin VB.Menu mnuMudaProximoTurno 
         Caption         =   "&Muda para Próximo Turno"
      End
      Begin VB.Menu mnuPontoFuncionario 
         Caption         =   "&Ponto de Funcionário"
      End
   End
   Begin VB.Menu mnuNFCe 
      Caption         =   "&NFCe"
      Begin VB.Menu mnuNFCeCupomComplementar 
         Caption         =   "NFCe Complementa&r"
      End
      Begin VB.Menu mnuCancelamentoNFCe 
         Caption         =   "Ca&ncelamento"
      End
      Begin VB.Menu mnuReimpressaoNFCe 
         Caption         =   "Reim&pressão NFC-e"
      End
   End
   Begin VB.Menu mnuSenha 
      Caption         =   "Senha"
   End
End
Attribute VB_Name = "Movimento_Nfce_Auto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flag_Movimento_Cupom_Fiscal As Integer
Dim lImpBematech As Boolean
Dim lImpSchalter As Boolean
Dim lImpMecaf As Boolean
Dim lImpQuick As Boolean
Dim lImpDaruma As Boolean
Dim lImpDarumaFW As Boolean
'Private lProcessadorNFCE As New NFCeDLL.ProcessaNFCEFronteira

'Dim lOpcao As String
Dim lFinalizaAutomatico As Boolean
Dim lGrupoCombustivel As Integer
Dim lBaixaAutomaticaNoEstoque As Boolean

Dim lHoraPegouNumeroCupom As Date
Dim lNumeroCupom As Long
Dim lNumeroUltimoCupom As Long
Dim lNumeroMovimentoCaixa As Long
Dim lData As Date
Dim lHora As Date
Dim lOrdem As Integer
Dim lEmpresa As Integer
Dim lPeriodo As Integer
Dim lDataCupom As Date
Dim l_vezes As Integer
Dim l_qtd_periodo As Integer
Dim lQtdBomba As Integer
Dim l_flag_cupom_fiscal As String
Dim lCodigoFiscal As String
Dim lTotalCupom As Currency
Dim lDescontoEspecial As Currency
Dim l_mensagem As String
Dim l_codigo_funcionario As Integer
Dim l_codigo_cliente As Integer
Dim l_nome_funcionario As String
Dim l_senha_funcionario As String
Dim lCupomDemonstracao As Boolean
Dim lInformaFormaPagamento As Boolean
Dim lDadosTCS As String
Dim x_tempo As Integer
Dim lOrigemAutomacao As Boolean
'Dim BemaRetorno As Integer
Dim lDescontoItemEmbutido As Currency
Dim lAcrescimoItemEmbutido As Currency
Dim lSQL As String
Dim lI As Integer
Dim lAutoBico As Integer
Dim lAutoQuantidade As Currency
Dim lAutoValorTotal As Currency
Dim lAutoHora As String
Dim lLegislacaoPermiteIssEcf  As Boolean
Dim lCodigoTcsEcf As Integer
Dim lContadorNaoFiscal As String
Dim lCodigoCartao As Integer
Dim lNumeroLancamentoCartao As Long
Dim lValorTotalUltimoCupom As Currency
Dim lNumeroPDV As Integer
Dim lSerieECF As String
Dim lTipoMovimento As Integer
Dim lValorUnitarioSemDesconto
Dim lValorTotalSemDesconto
Dim lCartaoAutorizacao As String
Dim lCartaoNSU As String
Dim lCartaoDataVencimento As String
Dim lTotalItem(0 To 30) As Currency
Dim lBloqueiaEstoque As Boolean
Dim lBloqueiaSubEstoque As Boolean
Dim lBloqueiaDesconto As Boolean

Dim lRestringeDescontoNivelUsuario As Boolean

Dim lRestringeVendaCredito As Integer
Dim lMarcaAutomacao As String
Dim lNomeArquivoAutomacaoIni As String
Dim lLoja As Boolean
Dim lCodigoBarra As Boolean
Dim lIlha As Integer
Dim lOrigemFocus As String
Dim lQtdMaxCombustivel As Currency
Dim lQtdMaxProduto As Currency
Dim lNotificacaoGic As Boolean
Dim lExisteMudancaHorarioVerao As Boolean
Dim lImprimeDocumentoVinculado As Boolean
Dim lPlacaLetra As String
Dim lPlacaNumero As Long
Dim lKMVeiculo As Long
Dim lEcfTruncamento As Boolean
Dim lEcfQtdCasasDecimais As Integer
Dim lCaixaIndividual As Boolean
Dim lGeraCaixaDinheiro As Boolean
Dim lGeraCaixaChequeAVista As Boolean
Dim lEcfInstalada As Boolean
'Dim lTestaReducaoZpendente As Boolean
Dim lNSU As Long
Dim lPortaRfid As Integer
Dim lExigeNCM As Boolean
Dim lCodigoVeiculo As Integer 'NEW 07/04
'Tef
Dim lQtdViasConfDiv As Integer
Dim lTEF As Boolean
Dim lRespostaTEF As Boolean
Dim lErroExtendido As String
Dim lAck As Integer
Dim lSt1 As Integer
Dim lSt2 As Integer
Dim lLinhasEntreCV As Integer
Dim lValorDescontoConcedido As Currency
Dim lPontuacaoConcedida As Currency
Dim lDescontoPostoAki As Boolean
Dim lDescontoCartaoEspecial As Boolean
Dim lIntegraDescontoCartaoCorreios As Boolean
Dim lIntegraDescontoCartaoEspecial As Boolean


Dim lSerieNFCe As String
Dim lNFCe_tPag As String
Dim lNFCe_vPag As Currency
Dim lNFCe_TpIntegra As Integer
Dim lNFCe_CNPJCartao As String
Dim lNFCe_tBand As String
Dim lNFCe_cAut As String
Dim lContadorAguarde As Integer

Dim lPOSImprimeNFCe As Boolean

Dim lTotalDescontoNivelPreco As Currency
Dim lTotalAcrescimoNivelPreco As Currency


Private AberturaCaixa As New cAberturaCaixa
Private Aliquota As New cAliquota
Private BaixaAbastecimento As New cBaixaAbastecimento
Private Bomba As New cBomba
Private CartaoAbastecimento As New cCartaoAbastecimento
Private CartaoCredito As New cCartaoCredito
Private CidadeIBGE As New cCidadeIBGE
Private Cliente As New cCliente
Private Combustivel As New cCombustivel
Private Configuracao As New cConfiguracao
Private ConfiguracaoDiversa As New cConfiguracaoDiversa
Private CerradoTef As CerradoComponenteTef
Private Credito As New cCredito
Private DuplicataReceber As New cDuplicataReceber
Private ECF As New cEcf
Private EncerranteAtual As New cEncerranteAtual
Private Estoque As New cEstoque
Private Funcionario As New cFuncionario
Private Grupo As New cGrupo
Private GrupoTipoMovimentoCaixa As New cGrupoTipoMovimentoCaixa
Private IntegracaoCaixa As New cIntegracaoCaixa
Private LiberacaoDigitacao As New cLiberacaoDigitacao
Private PeriodoTrocaOleo As New cPeriodoTrocaOleo
Private Produto As New cProduto
Private Produto2 As New cProduto
Private MovCaixaPista As New cMovimentoCaixaPista
Private MovCartaoCredito As New cMovimentoCartaoCredito
Private MovCupomFiscal As New cMovimentoCupomFiscal
Private MovCupomFiscalItem As New cMovimentoCupomFiscalItem
Private MovDescontoGrupoCliente As New cMovDescontoGrupoCliente
Private MovDescontoPersonalizado As New cMovDescontoPersonalizado
Private MovHorarioVerao As New cMovimentoHorarioVerao
Private MovMapaResumo As New cMovimentoMapaResumo
Private MovimentoLubrificante As New cMovimentoLubrificante
Private MovNotaAbastecimento As New cMovimentoNotaAbastecimento
Private MovimentoAbastecimento As New cMovimentoAbastecimento
Private MovimentoBomba As New cMovimentoBomba
Private MovimentoBombaEscritorio As New cMovimentoBomba
Private MovimentoCheque As New cMovimentoCheque
Private MovimentoChequeDevolvido As New cMovimentoChequeDevolvido
Private PercentualImposto As New cPercentualImposto
Private SubEstoque As New cSubEstoque
Private SolicitacaoFuncaoAutomacao As New cSolicitacaoFuncaoAutomacao
Private TaxaAdmCartaoCredito As New cTaxaAdmCartaoCredito
Private TicketCarDePara As New cTicketCarDePara
Private Usuario As New cUsuario
Private VeiculoCliente As New cVeiculoCliente 'NEW 04/07
Private MovSolicitacaoFuncaoNFe As New cMovSolicitacaoFuncaoNFe
Private lUtilizaPrecoAPrazo As Boolean
 

Dim rstAbastecimento As New ADODB.Recordset
Dim rsTabela As New ADODB.Recordset


Dim lAutomacaoFlag As Integer
Dim lAutomacaoFlagVendaAutomatica As Boolean
Dim lAutomacaoPorta As Integer
Dim lAutomacaoVelocidade As String
Dim lAutomacaoDtr As Boolean
Dim lAutomacaoRts As Boolean
Dim lAutomacaoBicoEmAcerto As Integer
Dim lAutomacaoDataEmAcerto As Date
Dim lAutomacaoHoraEmAcerto As Date
Dim lAutomacaoTempoAbastecimentoEmAcerto As String
Dim lAutomacaoFormaImpressao As String

Dim lAutomacaoStatusBico(0 To gQUANTIDADE_MAXIMA_BICO - 1) As Integer
Dim lAutomacaoCodigoProduto(0 To gQUANTIDADE_MAXIMA_BICO - 1) As Long
Dim lAutomacaoBico(0 To gQUANTIDADE_MAXIMA_BICO - 1) As Integer
Dim lAutomacaoData(0 To gQUANTIDADE_MAXIMA_BICO - 1) As Date
Dim lAutomacaoHora(0 To gQUANTIDADE_MAXIMA_BICO - 1) As Date
Dim lAutomacaoTempoAbastecimento(0 To gQUANTIDADE_MAXIMA_BICO - 1) As String
Dim lAutomacaoValorLitro(0 To gQUANTIDADE_MAXIMA_BICO - 1) As Currency
Dim lAutomacaoLitros(0 To gQUANTIDADE_MAXIMA_BICO - 1) As Currency
Dim lAutomacaoTotalAPagar(0 To gQUANTIDADE_MAXIMA_BICO - 1) As Currency

'Variáveis de Impressao
Dim lNomeArquivo As String
Dim lLocalImpressao As Integer
Dim lLinha As String

Dim rst As New ADODB.Recordset
Dim rst2  As New ADODB.Recordset

'--- CRIADO PARA NFCE ---
Const CODIGO_FISCAL_SUBSTITUICAO As String = "FF"
Const MODELO_NFCE As String = "65"
Const PROGRAMA_ORIGEM As String = "NFCE_AUTO"
Private MovDocEletronicoCabecalho As New cMovDocEletronicoCabecalho
Private MovDocEletronicoItem As New cMovDocEletronicoItem
Private CSTPisValidos As New Dictionary
Private CSTCofinsValidos As New Dictionary


Dim lUfEmpresa As String

Private Enum ETAPA_CONCLUIDA
    GRAVADO = 1
    PRE_PROCESSADO
    PROCESSANDO
    ERRO_Processamento
    NAO_IMPLEMENTADO5
    NAO_IMPLEMENTADO6
    DENEGADA
    REJEITADA
    AUTORIZADA
End Enum




Private Enum GERADOR_NFCE
    OOBJ = 1
    CERRADO
End Enum
Private Sub CalculaTotalDescontoAcrescimo(ByVal pData As Date, ByVal pPeriodo As Integer, ByVal pTipoCombustivel As String, pBico As Integer, ByVal pNumeroSubCaixa As Integer)
    Dim rs As New ADODB.Recordset
    
    On Error GoTo trata_erro

        lTotalDescontoNivelPreco = 0
        lTotalAcrescimoNivelPreco = 0
        
       If Not lUtilizaPrecoAPrazo Then
          Exit Sub
       End If


        If pNumeroSubCaixa <> 999 Then
            Exit Sub
        End If


        Dim xSQL As String
        xSQL = "SELECT SUM([Valor Desconto Nivel Preco]) as TotalDesconto, sum([Valor Acrescimo Nivel Preco]) as TotalAcrescimo FROM [dbo].[Movimento_Abastecimento]"
        xSQL = xSQL & " WHERE  data = " & preparaData(pData) '29/11/2018'
        xSQL = xSQL & " AND ([Valor Desconto Nivel Preco] > 0 or [Valor Acrescimo Nivel Preco] > 0)"
        xSQL = xSQL & " AND [Periodo] = " & pPeriodo
        xSQL = xSQL & " AND [Tipo de Combustivel] = " & preparaTexto(pTipoCombustivel)
        xSQL = xSQL & " AND Bico = " & pBico
        xSQL = xSQL & " AND [Valor Desconto Nivel Preco] IS NOT NULL"
        xSQL = xSQL & " AND [Valor Acrescimo Nivel Preco] IS NOT NULL"

        Set rs = Conectar.RsConexao(xSQL)
        
        If rs.RecordCount > 0 Then
           rs.MoveFirst
           
           If Not IsNull(rs("TotalDesconto").Value) Then
                lTotalDescontoNivelPreco = rs("TotalDesconto").Value
           End If
           
           If Not IsNull(rs("TotalAcrescimo").Value) Then
               lTotalAcrescimoNivelPreco = rs("TotalAcrescimo").Value
           End If
        End If
    
Exit Sub
trata_erro:
    MsgBox "Erro: CalculaTotalDescontoAcrescimo!" & Err.Description, vbInformation, "Erro de integridade"
Exit Sub
End Sub


Private Function CalculaImpostos(x_numero_cupom As Long, x_data As Date) As String
    Dim xBaseCalculo As Currency
    Dim xTotalCupom As Currency
    Dim xTotalImpostos As Currency
    Dim xPercentualImpostos As Currency
    Dim xDescontoCupom As Currency
    Dim xOrdem As Integer
    Dim xString As String
    
    'Call CriaLogCupom("CalculaImpostos: Fase 1=")
    CalculaImpostos = ""
    xBaseCalculo = 0
    xTotalCupom = 0
    xTotalImpostos = 0
    xPercentualImpostos = 0
    xDescontoCupom = 0
    xOrdem = 0
    
    'Call CriaLogCupom("CalculaImpostos: Fase 1 b lNumeroPDV=" & lNumeroPDV)
    'Call CriaLogCupom("CalculaImpostos: Fase 1 b x_numero_cupom=" & x_numero_cupom)
    'Call CriaLogCupom("CalculaImpostos: Fase 1 b x_data=" & x_data)
    'Call CriaLogCupom("CalculaImpostos: Fase 1 b xOrdem=" & xOrdem)
    Do Until MovCupomFiscal.LocalizarNumeroProximaOrdem(g_empresa, lNumeroPDV, x_numero_cupom, x_data, xOrdem) = False
        If MovCupomFiscal.ItemCancelado = False And MovCupomFiscal.CupomCancelado = False Then
            'Call CriaLogCupom("CalculaImpostos: Fase 2 a - MovCupomFiscal.CodigoProduto=" & MovCupomFiscal.CodigoProduto)
            If Produto.LocalizarCodigo(MovCupomFiscal.CodigoProduto) Then
                If LocalizarNCM(0, Produto.CodigoNCM) Then
                    xBaseCalculo = MovCupomFiscal.ValorTotal
                    'Call CriaLogCupom("CalculaImpostos: Fase 2 B - xBaseCalculo=" & xBaseCalculo)
                    'Call CriaLogCupom("CalculaImpostos: Fase 2 C - PercentualImposto.AliquotaNacional=" & PercentualImposto.AliquotaNacional)
                    'Call CriaLogCupom("CalculaImpostos: Fase 2 D - PercentualImposto.Codigo=" & PercentualImposto.Codigo)
                    'Call CriaLogCupom("CalculaImpostos: Fase 2 E - PercentualImposto.Ex=" & PercentualImposto.ex)
                    'Call CriaLogCupom("CalculaImpostos: Fase 2 F - Produto.CodigoNCM=" & Produto.CodigoNCM)
                    xTotalCupom = xTotalCupom + xBaseCalculo
                    xTotalImpostos = xTotalImpostos + (Round(xBaseCalculo * PercentualImposto.AliquotaNacional / 100, 2))
                Else
                    Call CriaLogCupom("CalculaImpostos: - NCM nao localizado. Produto.CodigoNCM=" & Produto.CodigoNCM)
                End If
            Else
                'Call CriaLogCupom("CalculaImpostos: Fase 2 z - Produto nao localizado. Codigo=" & MovCupomFiscal.CodigoProduto)
            End If
        End If
        xOrdem = MovCupomFiscal.Ordem
    Loop
    'Call CriaLogCupom("CalculaImpostos: Fase 2 xTotalCupom=" & xTotalCupom)
    'Call CriaLogCupom("CalculaImpostos: Fase 3 xTotalImpostos=" & xTotalImpostos)
    If xTotalCupom > 0 And xTotalImpostos > 0 Then
        xPercentualImpostos = Round(xTotalImpostos / xTotalCupom * 100, 2)
        'Call CriaLogCupom("CalculaImpostos: Fase 3a - xPercentualImpostos=" & xPercentualImpostos)
        xString = "Val.Aprox.Tributos R$ " & Format(xTotalImpostos, "###,##0.00") & "(" & Format(xPercentualImpostos, "##0.00") & "%) Fonte: IBPT"
        'Call CriaLogCupom("CalculaImpostos: xString=" & xString)
        If Len(xString) < 48 Then
            Do Until Len(xString) = 48
                xString = xString & " "
            Loop
        End If
        CalculaImpostos = Mid(xString, 1, 48)
    End If
    'Call CriaLogCupom("CalculaImpostos: Fase 4 xPercentualImpostos=" & xPercentualImpostos)
    Call CriaLogCupom("CalculaImpostos: CalculaImpostos=" & CalculaImpostos)
End Function
Private Sub CancelaCupom()
On Error GoTo trata_erro
    Call CriaLogMonitoraBug("CancelaCupom", "Movimento_NfceAuto", "Iniciou")
    frmDados.Enabled = False
    
    LimpaTela
    If l_flag_cupom_fiscal = "A" Then
        DespreparaDadosAdicionaisFechamento
        frm_fechamento_cupom.Enabled = True
        frm_fechamento_cupom.Top = 50 '100 '4300
        frm_fechamento_cupom.Left = 30 '70 '5970
        frm_fechamento_cupom.Height = 9500 '7480
        frm_fechamento_cupom.Width = 7850 '7460
        frm_fechamento_cupom.ZOrder 0
        cmdIniciaProcessoNFCe.Enabled = True
        btnDescontoCartaoEspecial.Enabled = False
        txt_cpf.Text = ""
        txt_nome_cliente.Text = ""
        txt_observacao.Text = ""
        txt_observacao_2.Text = ""
        txt_numero_nota_abastecimento.Text = ""
        If Val(l_codigo_cliente) > 0 Then
            If Val(Cliente.CGC) > 0 Then
                txt_cpf.Text = Mid(Cliente.CGC, 1, 2) + "." + Mid(Cliente.CGC, 3, 3) + "." + Mid(Cliente.CGC, 6, 3) + "/" + Mid(Cliente.CGC, 9, 4) + "-" + Mid(Cliente.CGC, 13, 2)
            ElseIf Cliente.CPF <> "" Then
                txt_cpf.Text = Mid(Cliente.CPF, 1, 3) + "." + Mid(Cliente.CPF, 4, 3) + "." + Mid(Cliente.CPF, 7, 3) + "-" + Mid(Cliente.CPF, 10, 2)
            End If
            txt_nome_cliente.Text = Cliente.RazaoSocial
            txt_observacao.Text = Cliente.Endereco
            txt_observacao_2.Text = Trim(Cliente.Bairro) & " - " & Trim(Cliente.Cidade)
        End If
        txt_valor_desconto.Text = Format(lValorDescontoConcedido, "###,##0.00")
        txt_placa.Text = ""
        txt_kilometragem.Text = ""
        txt_numero_cheque.Text = ""
        txt_telefone.Text = ""
        If lDescontoEspecial > 0 Then
            txt_valor_desconto.Text = Format(lDescontoEspecial, "###,##0.00")
        End If
        cbo_forma_pagamento.SetFocus
        lbl_valor_compra.Caption = Format(lTotalCupom - lValorDescontoConcedido, "###,##0.00")
        txt_valor_recebido.Text = Format(lTotalCupom - lValorDescontoConcedido, "###,##0.00")
        lbl_valor_troco.Caption = Format(0, "0.00")
        txt_valor_recebido.SelStart = 0
        txt_valor_recebido.SelLength = Len(txt_valor_recebido.Text)
        If Not lInformaFormaPagamento Then
            If l_codigo_cliente = 0 Then
                DespreparaDadosAdicionaisFechamento
                cbo_forma_pagamento.ListIndex = 0
                cbo_forma_pagamento_LostFocus
                Call CriaLogMonitoraBug("CancelaCupom", "Movimento_NfceAuto", "Chamou IniciaProcessamentoNFCe")
                IniciaProcessamentoNFCe
            Else
                DespreparaDadosAdicionaisFechamento
                cbo_forma_pagamento.ListIndex = 4
                cbo_forma_pagamento_LostFocus
                Call CriaLogMonitoraBug("CancelaCupom", "Movimento_NfceAuto", "Chamou IniciaProcessamentoNFCe")
                IniciaProcessamentoNFCe
            End If
        Else
            cbo_forma_pagamento.ListIndex = 0
        End If
        If Val(l_codigo_cliente) > 0 Then
            cbo_forma_pagamento.ListIndex = 4
            txt_numero_nota_abastecimento.SetFocus
            If l_codigo_cliente = 47 Then
                txt_valor_recebido.SetFocus
            End If
        End If
    End If
    Call BuscaRegistro(lNumeroCupom, lData, lOrdem)
    Call CriaLogMonitoraBug("CancelaCupom", "Movimento_NfceAuto", "Finalizou")
    Exit Sub

trata_erro:
    Call CriaLogMonitoraBug("CancelaCupom", "Movimento_NfceAuto", "Deu erro: " & Err.Description)
    Call CriaLogCupom("Erro CancelaCupom: Erro=" & Err.Number & " - " & Err.Description)
End Sub
Private Function CancelamentoCupomFiscal() As Boolean
    Dim NumeroArquivo As Integer
    Dim x_excluiu As Boolean
    Dim rs As New ADODB.Recordset
    
    On Error GoTo FileError
    
    CancelamentoCupomFiscal = False
    'Localiza o último cupom fiscal
    If MovCupomFiscal.LocalizarUltimo(g_empresa, lNumeroPDV) Then
        If MovCupomFiscal.CupomCancelado = True Then
            Call GravaAuditoria(1, Me.name, 25, "Cancelamento abortado. Já está cancelado o ECF:" & MovCupomFiscal.NumeroCupom)
            MsgBox "Não será possível continuar o cancelamento!" & Chr(10) & "Este cupom já encontra-se cancelado.", vbInformation, "Cancelamento Negado!"
            Exit Function
        End If
    Else
        MsgBox "Não foi possível localizar o último Cupom Fiscal!", vbCritical, "Erro de Integridade!"
        Call GravaAuditoria(1, Me.name, 25, "Não foi possível localizar o último ECF. lNumeroPDV=" & lNumeroPDV)
        Exit Function
    End If
    
    lSQL = "SELECT * FROM Movimento_Cupom_Fiscal"
    lSQL = lSQL & " WHERE Data = " & preparaData(MovCupomFiscal.Data)
    lSQL = lSQL & "   AND [Numero do Cupom] = " & MovCupomFiscal.NumeroCupom
    lSQL = lSQL & "   AND [Cupom Cancelado] = " & preparaBooleano(False)
    lSQL = lSQL & "   AND [Item Cancelado] = " & preparaBooleano(False)
    lSQL = lSQL & "   AND [Codigo da ECF] = " & lNumeroPDV
    lSQL = lSQL & " ORDER BY Ordem"
    Call CriaLogCupom("CancelamentoCupomFiscal: Investigando...lSQL=" & lSQL)
    Set rs = Conectar.RsConexao(lSQL)
    
    x_excluiu = False
    
    'Cancela ECF na Impressora
    If lImpBematech Then
        If Not Testa_ImpressoraCF Then
            NumeroArquivo = 99999
        End If
'            Call CriaLogCupom("Bematech_FI_CancelaCupom")
'            BemaRetorno = Bematech_FI_CancelaCupom
'            Call CriaLogCupom("Bematech_FI_CancelaCupom - BemaRetorno=" & BemaRetorno)
    ElseIf lImpDaruma Then
'            Call CriaLogCupom("Daruma_FI_CancelaCupom")
'            BemaRetorno = Daruma_FI_CancelaCupom
'            Call CriaLogCupom("Daruma_FI_CancelaCupom - BemaRetorno=" & BemaRetorno)
        '30/03/16
    ElseIf lImpQuick Then
'            Call CriaLogCupom("CancelaCupom")
'            BemaRetorno = EcfQuickCancelaCupom
'            Call CriaLogCupom("CancelaCupom - BemaRetorno=" & BemaRetorno) '
    End If
    
    'If BemaRetorno <> 1 Then
    '    Call GravaAuditoria(1, Me.name, 25, "Não permitido por tempo excedido. ECF:" & MovCupomFiscal.NumeroCupom)
    '    MsgBox "Cancelamento do último cupom não permitido(1)." & Chr(10) & "Tempo excedido!", vbInformation, "Cancelamento Não Permitido!"
    '    Exit Function
    'End If
    
    'Cancela ECF no sistema
    MovCupomFiscal.CupomCancelado = True
    Call CriaLogCupom("CancelamentoCupomFiscal: Investigando 1... rs(Numero do Cupom).Value=" & rs("Numero do Cupom").Value & " - rs(Data).Value=" & rs("Data").Value & " - rs(Ordem).Value=" & rs("Ordem").Value)
    If MovCupomFiscal.LocalizarCodigo(g_empresa, lNumeroPDV, rs("Numero do Cupom").Value, rs("Data").Value, rs("Ordem").Value) Then
        Call CriaLogCupom("CancelamentoCupomFiscal: Investigando 2... lNumeroPDV=" & lNumeroPDV & " - MovCupomFiscal.NumeroCupom=" & MovCupomFiscal.NumeroCupom & " - MovCupomFiscal.Data=" & MovCupomFiscal.Data)
        If MovCupomFiscal.CancelaCupom(g_empresa, lNumeroPDV, MovCupomFiscal.NumeroCupom, MovCupomFiscal.Data) Then
            If MovCupomFiscalItem.CancelaCupom(g_empresa, lNumeroPDV, MovCupomFiscal.Data, MovCupomFiscal.NumeroCupom) Then
                If Not rs.EOF Then
                    rs.MoveFirst
                    Do Until rs.EOF
                        Call CriaLogCupom("CancelamentoCupomFiscal: Investigando 3... rs(Numero do Cupom).Value=" & rs("Numero do Cupom").Value & " - rs(Data).Value=" & rs("Data").Value & " - rs(Ordem).Value=" & rs("Ordem").Value)
                        If MovCupomFiscal.LocalizarCodigo(g_empresa, lNumeroPDV, rs("Numero do Cupom").Value, rs("Data").Value, rs("Ordem").Value) Then
                            CancelamentoCupomFiscal = True
                            Call GravaAuditoria(1, Me.name, 25, "Cancelado o ECF:" & rs("Numero do Cupom").Value & " Ítem:" & rs("Ordem").Value)
                            If Produto.LocalizarCodigo(rs("Codigo do Produto").Value) Then
                                If Produto.CodigoGrupo = lGrupoCombustivel Then
                                    If g_automacao Then
                                        If Not MovimentoAbastecimento.VoltaEcfCancelado(g_empresa, MovCupomFiscal.Data, MovCupomFiscal.NumeroCupom) Then
                                            MsgBox "Não foi possível voltar o abastecimento.", vbCritical, "Erro de Integridade!"
                                            Call GravaAuditoria(1, Me.name, 25, "Erro ao extornar abastecimento do ECF:" & rs("Numero do Cupom").Value & " Ítem:" & rs("Ordem").Value)
                                        End If
                                    End If
                                Else
                                    If Configuracao.ECFBaixaEstoque = True Then
                                        Call SubtraiVendaProduto
                                    End If
                                End If
                            Else
                                MsgBox "Não foi possível localizar o produto:" & rs("Codigo do Produto").Value, vbCritical, "Erro de Integridade!"
                                Call GravaAuditoria(1, Me.name, 25, "Erro ao localizar o produto:" & rs("Codigo do Produto").Value)
                            End If
                            If lBaixaAutomaticaNoEstoque = True Then
                                Call AdicionaEstoque(rs("Codigo do Produto").Value, rs("Quantidade").Value, rs("Tipo do Movimento").Value)
                            End If
                            If MovCupomFiscal.CodigoCliente > 0 Then
                                ExcluiNotaAbastecimento
                            End If
                        Else
                            MsgBox "Não foi possível localizar o ítem do cupom no sistema.", vbCritical, "Erro de Integridade!"
                            Call GravaAuditoria(1, Me.name, 25, "Erro ao localizar o ítem:" & rs("Ordem").Value & " do ECF:" & rs("Numero do Cupom").Value)
                        End If
                        rs.MoveNext
                    Loop
                End If
            Else
                MsgBox "Não foi possível cancelar os ítens do cupom no sistema.", vbCritical, "Erro de Integridade!"
                Call GravaAuditoria(1, Me.name, 25, "Erro ao cancelar no sistema os ítens do ECF:" & MovCupomFiscal.NumeroCupom)
            End If
        Else
            MsgBox "Não foi possível cancelar o cupom fiscal no sistema.", vbCritical, "Erro de Integridade!"
            Call GravaAuditoria(1, Me.name, 25, "Erro ao cancelar no sistema o ECF:" & MovCupomFiscal.NumeroCupom)
        End If
    Else
        MsgBox "Não foi possível localizar o Cupom Fiscal N:" & rs("Numero do Cupom").Value, vbCritical, "Erro de Integridade!"
        Call GravaAuditoria(1, Me.name, 25, "Não foi possível localizar o ECF:" & rs("Numero do Cupom").Value)
    End If
    Exit Function

FileError:
    Call CriaLogCupom("Erro CancelamentoCupomFiscal: Erro=" & Err.Number & " - " & Err.Description)
    Call GravaAuditoria(1, Me.name, 25, "CancelamentoCupomFiscal: Erro inesperado...")
    Exit Function
End Function
Private Function CancelamentoCupomFiscalItem() As Boolean
    Dim NumeroArquivo As Integer
    
    On Error GoTo FileError
    
    CancelamentoCupomFiscalItem = False
    If MovCupomFiscal.LocalizarCodigo(g_empresa, lNumeroPDV, lNumeroCupom, lData, lOrdem - 1) Then
        If MovCupomFiscal.CupomCancelado = True Then
            Call GravaAuditoria(1, Me.name, 25, "Cancelamento abortado. Já está cancelado o ECF:" & lNumeroCupom)
            MsgBox "Não será possível continuar o cancelamento!" & Chr(10) & "Este cupom já encontra-se cancelado.", vbInformation, "Cancelamento Negado!"
            Exit Function
        ElseIf MovCupomFiscal.ItemCancelado = True Then
            Call GravaAuditoria(1, Me.name, 25, "Cancelamento abortado. Já está cancelado o ECF:" & lNumeroCupom & " Ítem:" & MovCupomFiscal.Ordem)
            MsgBox "Não será possível continuar o cancelamento!" & Chr(10) & "Este ítem de cupom já encontra-se cancelado.", vbInformation, "Cancelamento Negado!"
            Exit Function
        End If
    Else
        Call GravaAuditoria(1, Me.name, 25, "Não foi possível localizar o ECF:" & lNumeroCupom & " Ítem:" & lOrdem - 1)
        MsgBox "Não foi possível localizar o cupom fiscal para cancelar!", vbCritical, "Erro de Integridade!"
        Exit Function
    End If
    
    'Cancela Ítem na Impressora
    If lImpBematech Then
        If Not Testa_ImpressoraCF Then
            NumeroArquivo = 99999
        End If
'            Call CriaLogCupom("Bematech_FI_CancelaItemAnterior")
'            BemaRetorno = Bematech_FI_CancelaItemAnterior()
'            Call CriaLogCupom("Bematech_FI_CancelaItemAnterior - BemaRetorno=" & BemaRetorno)
    ElseIf lImpDaruma Then
'            Call CriaLogCupom("Daruma_FI_CancelaItemAnterior")
'            BemaRetorno = Daruma_FI_CancelaItemAnterior
'            Call CriaLogCupom("Daruma_FI_CancelaItemAnterior - BemaRetorno=" & BemaRetorno)
    'End If
    '28/03/16
    ElseIf lImpQuick Then
        If MovCupomFiscal.LocalizarCodigo(g_empresa, lNumeroPDV, lNumeroCupom, lData, lOrdem - 1) Then
'                Call CriaLogCupom("CancelaItemFiscal")
'                BemaRetorno = EcfQuickCancelaItemFiscal(MovCupomFiscal.Ordem)
'                Call CriaLogCupom("CancelaItemFiscal - BemaRetorno=" & BemaRetorno)
        End If
    End If
    '
    'If BemaRetorno <> 1 Then
    '    Call GravaAuditoria(1, Me.name, 25, "Não permitido por tempo excedido. ECF:" & lNumeroCupom & " Ordem:" & MovCupomFiscal.Ordem)
    '    MsgBox "Cancelamento do último cupom não permitido(2)." & Chr(10) & "Tempo excedido!", vbInformation, "Cancelamento Não Permitido!"
    '    Exit Function
    'End If
        
    'Cancela Ítem no sistema
    MovCupomFiscal.ItemCancelado = True
    If MovCupomFiscal.CancelaItemCupom(g_empresa, lNumeroPDV, lNumeroCupom, lData, lOrdem - 1) Then
        If MovCupomFiscalItem.CancelaItem(g_empresa, lNumeroPDV, lData, lNumeroCupom, lOrdem - 1) Then
            CancelamentoCupomFiscalItem = True
            Call GravaAuditoria(1, Me.name, 25, "Cancelado o ítem de ECF:" & lNumeroCupom & " Ítem:" & lOrdem - 1)
        Else
            MsgBox "Não foi possível cancelar o ítem do cupom.", vbCritical, "Erro de Integridade!"
            Call GravaAuditoria(1, Me.name, 25, "Erro ao cancelar no sistema ítem do ECF:" & lNumeroCupom & " Ordem=" & lOrdem - 1)
        End If
        If lBaixaAutomaticaNoEstoque = True Then
            Call AdicionaEstoque(MovCupomFiscal.CodigoProduto, MovCupomFiscal.Quantidade, MovCupomFiscal.TipoSubEstoque)
        End If
        If Produto.LocalizarCodigo(MovCupomFiscal.CodigoProduto) Then
            If Produto.CodigoGrupo = lGrupoCombustivel Then
                If g_automacao Then
                    If Not MovimentoAbastecimento.VoltaEcfCancelado(g_empresa, MovCupomFiscal.Data, MovCupomFiscal.NumeroCupom) Then
                        MsgBox "Não foi possível voltar o abastecimento.", vbCritical, "Erro de Integridade!"
                        Call GravaAuditoria(1, Me.name, 25, "Erro ao extornar abastecimento do ECF:" & MovCupomFiscal.NumeroCupom & " Ítem:" & MovCupomFiscal.Ordem)
                    End If
                End If
            Else
                If Configuracao.ECFBaixaEstoque = True Then
                    Call SubtraiVendaProduto
                End If
            End If
        Else
            MsgBox "Não foi possível localizar o produto:" & MovCupomFiscal.CodigoProduto, vbCritical, "Erro de Integridade!"
            Call GravaAuditoria(1, Me.name, 25, "Erro ao localizar o produto:" & MovCupomFiscal.CodigoProduto)
        End If
        If MovCupomFiscal.CodigoCliente > 0 Then
            ExcluiNotaAbastecimento
        End If
    Else
        MsgBox "Não foi possível cancelar o ítem do cupom fiscal no sistema.", vbCritical, "Erro de Integridade!"
        Call GravaAuditoria(1, Me.name, 25, "Erro ao cancelar no sistema do ECF:" & lNumeroCupom & " Ordem=" & lOrdem - 1)
    End If
    Exit Function

FileError:
    Call CriaLogCupom("Erro CancelamentoCupomFiscalItem: Erro=" & Err.Number & " - " & Err.Description)
    Call GravaAuditoria(1, Me.name, 25, "CancelamentoCupomFiscalItem: Erro inesperado...")
    Exit Function
End Function
Private Function CasaDecimalZerada(ByVal pValor As String)
    Dim xStringDecimal As Variant
    
    CasaDecimalZerada = False
    xStringDecimal = Split(Format(fValidaValor(pValor), "#########0.0000"), ",")
    If CInt(xStringDecimal(1)) = 0 Then
        CasaDecimalZerada = True
    End If
End Function
Private Function ChamaCartaoDesconto(ByVal pTipoDesconto As String, ByVal pNumeroAutorizacaoPostoAki As String, ByVal pTrocaOleo As Boolean, ByVal pPontuacao As Boolean) As Currency
    Dim xString As String
    Dim xTextoParaComprovante As String
    Dim xDadosProdutos As String
    Dim xImprimeTef As String
    
    Dim xNomeFuncionarioDesconto As String
    Dim xCodigoFuncionarioDesconto As Integer
    Dim xFrmSelecionaFuncionario As New frmSelecionaFuncionario
    Dim xDeveSelecionarFuncionario As Boolean
    
    xNomeFuncionarioDesconto = l_nome_funcionario
    xCodigoFuncionarioDesconto = l_codigo_funcionario
    xDeveSelecionarFuncionario = False
    
    ChamaCartaoDesconto = 0
    Call CriaLogECF(Date & " " & Time & " TEF: N.Cupom=" & lNumeroCupom & " - Valor=" & txt_valor_recebido.Text & " - Forma Pg.=" & cbo_forma_pagamento.Text & " - Pontuação: " & pPontuacao)
    gNumeroControleSolicitacao = Configuracao.ProximaSolicitacaoTEF(g_empresa)
    lRespostaTEF = False
    'Set CerradoTef = Nothing
    Set CerradoTef = New CerradoComponenteTef
    
    If pTipoDesconto = "POSTOAKI" Then
        If ConfiguracaoDiversa.LocalizarCodigo(1, "Informar funcionário para PostoAki") Then
            xDeveSelecionarFuncionario = ConfiguracaoDiversa.Verdadeiro
        End If
    End If
    
    If xDeveSelecionarFuncionario Then
        xFrmSelecionaFuncionario.Show vbModal
        
        If xFrmSelecionaFuncionario.gCodigoFuncionarioSelecionado = 0 Then
            lValorDescontoConcedido = 0
            lPontuacaoConcedida = 0
            ChamaCartaoDesconto = 0
            MsgBox "Para aplicar desconto o funcionário deve ser selecionado.", vbInformation, "Funcionário não selecionado"
            Exit Function
        Else
            xNomeFuncionarioDesconto = xFrmSelecionaFuncionario.gNomeFuncionarioSelecionado
            xCodigoFuncionarioDesconto = xFrmSelecionaFuncionario.gCodigoFuncionarioSelecionado
        End If
    End If
    
    'Prepara Texto para sair no comprovante de venda
    'ALEX 01/11/2017
    'xTextoParaComprovante = "Funcionario:" + Format(l_codigo_funcionario, "000") + " " + Mid(l_nome_funcionario, 1, 15)
    xTextoParaComprovante = "Funcionario:" + Format(xCodigoFuncionarioDesconto, "000") + " " + Mid(xNomeFuncionarioDesconto, 1, 15)
    If Len(xTextoParaComprovante) < 48 Then
        Do While Len(xTextoParaComprovante) <= 48
            xTextoParaComprovante = xTextoParaComprovante & " "
        Loop
    End If
    xTextoParaComprovante = String(48, "-") & xTextoParaComprovante & String(48, "-")
    xDadosProdutos = "" & vbCrLf & PreparaDadosProdutos
    'aqui 08/04/2016
    'lValorDescontoConcedido = CerradoTef.SolicitacaoDesconto("NFCe", gNumeroControleSolicitacao, lNumeroCupom, txt_valor_recebido.Text, gQtdViasTEF, xDadosProdutos, lLinhasEntreCV, xTextoParaComprovante)
    'Tasso 16/11/2016
    'lValorDescontoConcedido = CerradoTef.SolicitacaoDesconto("NFCe", gNumeroControleSolicitacao, lNumeroCupom, txt_valor_recebido.Text, gQtdViasTEF, xDadosProdutos, lLinhasEntreCV, xTextoParaComprovante, pTipoDesconto, pNumeroAutorizacaoPostoAki, l_codigo_funcionario, l_nome_funcionario)
    'ALEX 01/11/2017
    'lValorDescontoConcedido = CerradoTef.SolicitacaoDesconto("NFCe", gNumeroControleSolicitacao, lNumeroCupom, lbl_valor_compra.Caption, gQtdViasTEF, xDadosProdutos, lLinhasEntreCV, xTextoParaComprovante, pTipoDesconto, pNumeroAutorizacaoPostoAki, l_codigo_funcionario, l_nome_funcionario, pTrocaOleo)
    If pPontuacao = True Then
        lPontuacaoConcedida = CerradoTef.SolicitacaoDesconto("NFCe", gNumeroControleSolicitacao, lNumeroCupom, lbl_valor_compra.Caption, gQtdViasTEF, xDadosProdutos, lLinhasEntreCV, xTextoParaComprovante, pTipoDesconto, pNumeroAutorizacaoPostoAki, xCodigoFuncionarioDesconto, xNomeFuncionarioDesconto, pTrocaOleo, pPontuacao)
        ChamaCartaoDesconto = lPontuacaoConcedida
        Call CriaLogCupom("Cartão Fidelidade/Pontuação. - lPontuacaoConcedida=" & lPontuacaoConcedida)
    Else
        lValorDescontoConcedido = CerradoTef.SolicitacaoDesconto("NFCe", gNumeroControleSolicitacao, lNumeroCupom, lbl_valor_compra.Caption, gQtdViasTEF, xDadosProdutos, lLinhasEntreCV, xTextoParaComprovante, pTipoDesconto, pNumeroAutorizacaoPostoAki, xCodigoFuncionarioDesconto, xNomeFuncionarioDesconto, pTrocaOleo, pPontuacao)
        ChamaCartaoDesconto = lValorDescontoConcedido
        Call CriaLogCupom("Cartão Fidelidade/Desconto. - lValorDescontoConcedido=" & lValorDescontoConcedido)
    End If
    
    lDescontoPostoAki = False
    If pTipoDesconto = "POSTOAKI" Then
        If pPontuacao = False And lValorDescontoConcedido > 0 Then
            lDescontoPostoAki = True
            'PREENCHER COMBO FORMA DE PAGAMENTO SOMENTE COM DINHEIRO E CARTÃO AQUI - ALEXPOSTOAKI
            Call PreencheCboFormaPagamentoParaPostoAki
        ElseIf pPontuacao = True And lPontuacaoConcedida > 0 Then
            lDescontoPostoAki = True
        End If
    End If
    
    'ChamaCartaoDesconto = lValorDescontoConcedido
    Set CerradoTef = Nothing
    If lValorDescontoConcedido > 0 Then
        txt_valor_desconto.Text = Format(lValorDescontoConcedido, "###,##0.00")
        txt_valor_desconto.Enabled = False
        lbl_valor_compra.Caption = Format(lTotalCupom - fValidaValor(txt_valor_desconto.Text), "###,##0.00")
        txt_valor_recebido.Text = Format(lTotalCupom - fValidaValor(txt_valor_desconto.Text), "###,##0.00")
        ' ESTUDAR SE PRECISA: Melhorar CerradoTef.SolicitacaoDesconto, para trazer o NOME do usuário do Cartão
        If Not IncluiMovimentoCaixa(lDataCupom, lPeriodo, True, "DescontoCartaoFidelidade", lValorDescontoConcedido, "", "Cartão Fidelidade") Then
            Call CriaLogCupom("Erro ChamaCartaoDesconto:Desconto/Acréscimo não integrada no caixa. Cliente=" & MovCupomFiscal.CodigoCliente)
            MsgBox "Não foi possível integrar Desconto/Acréscimo no caixa!", vbInformation, "Erro de Integridade!"
        End If
        If lImpBematech Then
            Call CriaLogCupom("Cartão Fidelidade/Desconto - lValorDescontoConcedido=" & lValorDescontoConcedido)
            'Desconto para o Cupom Fiscal
            xString = Mid(Format(lValorDescontoConcedido, "000000000000.00"), 1, 12) + Mid(Format(lValorDescontoConcedido, "000000000000.00"), 14, 2)
'                Call CriaLogCupom("Bematech_FI_IniciaFechamentoCupom('D', '$', xString) xString=" & xString)
'                BemaRetorno = Bematech_FI_IniciaFechamentoCupom("D", "$", xString)
'                Call CriaLogCupom("Bematech_FI_IniciaFechamentoCupom - BemaRetorno=" & BemaRetorno)
        End If
    End If
End Function
Private Sub ChamaCalcLitros()
    'g_valor = fValidaValor4(txt_valor_unitario)
    'calc_litro.Show 1
    'txt_quantidade = Format(RetiraGString(1), "###,##0.000")
    'txt_valor_total = Format(RetiraGString(2), "###,##0.00")
    'GravaItem
End Sub
Private Sub SaiHorarioVerao()
    Timer2.Enabled = False
    Timer2.Interval = 0
    TimerAutomacao.Enabled = False
    TimerAutomacao.Interval = 0
    
    If txt_funcionario_ponto.Enabled = True Then
        txt_funcionario_ponto.SetFocus
    Else
        txt_produto.SetFocus
    End If
    Me.lbl_mensagem.Alignment = 2
    Me.lbl_mensagem.Caption = "Aguarde! Sistema em Processo para Sair do Horário de Verão!"
    txt_cupom_fiscal.Visible = False
    frm_fila_bico.Visible = False
    frmDados.Visible = False
    frm_ponto.Visible = False
    frm_fechamento_cupom.Visible = False
    mnuFuncaoADM.Enabled = False
    mnuPontoFuncionario.Enabled = False
    mnuTCS.Enabled = False
    mnuSenha.Enabled = False
    mnuConsultaCheque.Enabled = False
    mnuVisualizaVenda.Enabled = False
    mnuCancelaCartao.Enabled = False
    cmdIniciaProcessoFinalizacaoNFCe.Visible = False
    mnuFechamentoCaixa.Enabled = False
    
    frmDescarregar.Caption = "Programação de Horário de Verão"
    frmDescarregar.Enabled = True
    frmDescarregar.Visible = True
    frmDescarregar.Top = 20
    frmDescarregar.Left = 40
    frmDescarregar.Width = 11800
    frmDescarregar.Height = 4000
    Label14.Top = 1500
    Label14.Left = 2500
    Label14.AutoSize = True
    Label14.Caption = "Favor não desligar o Computador!"
    Label14.FontSize = 20
    cboBico.Visible = False
    Label15.Visible = False
    txtQuantidadeDescarregamento.Visible = False
    Label16.Visible = False
    chkDesconto.Visible = False
    cmdOkDescarregar.Visible = False
    cmdCancelarDescarregar.Visible = False
    
    Call GravaAuditoria(1, Me.name, 26, "Imprime Redução Z Automaticamente p/ sair do Horário de Verão")
    Do Until Date > CDate("21/03/2006")
        DoEvents
        Me.lbl_mensagem.Caption = "O sistema está aguardando o momento oportuno para voltar a data da Impressora Fiscal."
    Loop
    
    Do Until Time >= "01:05:00"
        DoEvents
        Me.lbl_mensagem.Caption = "O sistema está aguardando o momento oportuno para voltar a data da Impressora Fiscal!"
    Loop
'    Call CriaLogCupom("Bematech_FI_ProgramaHorarioVerao")
'    BemaRetorno = Bematech_FI_ProgramaHorarioVerao
'    Call CriaLogCupom("Bematech_FI_ProgramaHorarioVerao - BemaRetorno=" & BemaRetorno)
    MsgBox "Processo de Sair do Horário de Verão Concluído!" & Chr(10) & "Após teclar enter o sistema será fechado!" & Chr(10) & "Basta chamar o sistema Novamente e usar normalmente.", vbInformation, "Concluído!"
    End
End Sub
Private Sub SubtraiEstoque(ByVal pCodigoProduto As Long, ByVal pQuantidade As Currency, ByVal pTipoSubEstoque As Integer)
On Error GoTo trata_erro
    
    If Estoque.LocalizarCodigo(g_empresa, pCodigoProduto) Then
        Estoque.Quantidade = Estoque.Quantidade - pQuantidade
        If Estoque.Alterar(g_empresa, pCodigoProduto) Then
            If SubEstoque.AlterarQuantidade(g_empresa, pCodigoProduto, pTipoSubEstoque, pQuantidade, False) Then
            Else
                Call CriaLogCupom("Erro SubtraiEstoque:Sub-Estoque não alterado. Produto=" & pCodigoProduto & " Quantidade=" & pQuantidade & " SubEst=" & pTipoSubEstoque)
                MsgBox "Não foi possível alterar o sub-estoque!", vbInformation, "Erro de Integridade!"
            End If
        Else
            Call CriaLogCupom("Erro SubtraiEstoque:Estoque não alterado. Produto=" & pCodigoProduto & " Quantidade=" & pQuantidade & " SubEst=" & pTipoSubEstoque)
            MsgBox "Não foi possível alterar o estoque!", vbInformation, "Erro de Integridade!"
        End If
    Else
        Call CriaLogCupom("Erro SubtraiEstoque:Estoque não cadastrado. Produto=" & pCodigoProduto & " Quantidade=" & pQuantidade & " SubEst=" & pTipoSubEstoque)
        MsgBox "Estoque não cadastrado.", vbInformation, "Erro de Verificação!"
    End If
    Exit Sub

trata_erro:
    Call CriaLogCupom("Erro SubtraiEstoque:Desconhecido. Produto=" & pCodigoProduto & " Quantidade=" & pQuantidade & " SubEst=" & pTipoSubEstoque)
    Call CriaLogCupom("Erro SubtraiEstoque: Erro=" & Err.Number & " - " & Err.Description)
    MsgBox Err.Number & " - " & Err.Description
End Sub
Private Sub SubtraiVendaProduto()
On Error GoTo trata_erro
    
    If Not IntegracaoCaixa.LocalizarNome(g_empresa, "VENDA DE LUBRIFICANTES") Then
        MsgBox "Não será possível integrar com o caixa!", vbCritical, "Erro de Integridade!"
        Call GravaAuditoria(1, Me.name, 25, "Não será integrado no caixa o extorno de produto.")
        Exit Sub
    End If
        
    If ExcluiMovimentoCaixa("VENDA DE LUBRIFICANTES") Then
        If MovimentoLubrificante.LocalizarCodigo(g_empresa, MovCupomFiscal.Data, MovCupomFiscal.Periodo, lIlha, lTipoMovimento, MovCupomFiscal.TipoSubEstoque, MovCupomFiscal.CodigoProduto, MovCupomFiscal.operador) Then
            MovimentoLubrificante.Quantidade = MovimentoLubrificante.Quantidade - MovCupomFiscal.Quantidade
            MovimentoLubrificante.ValorTotal = MovimentoLubrificante.ValorTotal - MovCupomFiscal.ValorTotal
            If MovimentoLubrificante.Quantidade = 0 Then
                If MovimentoLubrificante.Excluir(g_empresa, MovCupomFiscal.Data, MovCupomFiscal.Periodo, lIlha, lTipoMovimento, MovCupomFiscal.TipoSubEstoque, MovCupomFiscal.CodigoProduto, MovCupomFiscal.operador) Then
                Else
                    Call GravaAuditoria(1, Me.name, 25, "Não excluiu venda de produto:" & MovCupomFiscal.CodigoProduto)
                    Call CriaLogCupom(Time & "SubtraiVendaProduto: Não excluiu venda de produto:" & MovCupomFiscal.CodigoProduto & " Data:" & MovCupomFiscal.Data & " Per:" & MovCupomFiscal.Periodo & " Ilha:" & MovCupomFiscal.CodigoECF & " Tipo Mov:" & MovCupomFiscal.TipoMovimento & " SubEst:" & MovCupomFiscal.TipoSubEstoque & " Prod:" & MovCupomFiscal.CodigoProduto & " Operador:" & MovCupomFiscal.operador)
                    MsgBox "Não foi possível excluir venda de produtos.", vbCritical, "Erro de Integridade!"
                End If
            Else
                If MovimentoLubrificante.Alterar(g_empresa, MovCupomFiscal.Data, MovCupomFiscal.Periodo, lIlha, lTipoMovimento, MovCupomFiscal.TipoSubEstoque, MovCupomFiscal.CodigoProduto, MovCupomFiscal.operador) Then
                Else
                    Call GravaAuditoria(1, Me.name, 25, "Não alterou venda de produto:" & MovCupomFiscal.CodigoProduto)
                    Call CriaLogCupom(Time & "SubtraiVendaProduto: Não alterou venda de produto:" & MovCupomFiscal.CodigoProduto & " Data:" & MovCupomFiscal.Data & " Per:" & MovCupomFiscal.Periodo & " Ilha:" & MovCupomFiscal.CodigoECF & " Tipo Mov:" & MovCupomFiscal.TipoMovimento & " SubEst:" & MovCupomFiscal.TipoSubEstoque & " Prod:" & MovCupomFiscal.CodigoProduto & " Operador:" & MovCupomFiscal.operador)
                    MsgBox "Não foi possível alterar venda de produtos.", vbCritical, "Erro de Integridade!"
                End If
            End If
        Else
            Call GravaAuditoria(1, Me.name, 25, "Não localizou venda de produto:" & MovCupomFiscal.CodigoProduto)
            Call CriaLogCupom(Time & "SubtraiVendaProduto: Não localizou venda de produto:" & MovCupomFiscal.CodigoProduto & " Data:" & MovCupomFiscal.Data & " Per:" & MovCupomFiscal.Periodo & " Ilha:" & MovCupomFiscal.CodigoECF & " Tipo Mov:" & MovCupomFiscal.TipoMovimento & " SubEst:" & MovCupomFiscal.TipoSubEstoque & " Prod:" & MovCupomFiscal.CodigoProduto & " Operador:" & MovCupomFiscal.operador)
            MsgBox "Não foi possível localizar venda de produtos.", vbCritical, "Erro de Integridade!"
        End If
    Else
        Call GravaAuditoria(1, Me.name, 25, "Não foi possível estornar venda de produto no caixa.")
        MsgBox "Não foi possível estornar no caixa!", vbCritical, "Erro de Integridade!"
    End If
    Exit Sub

trata_erro:
    Call CriaLogCupom("Erro SubtraiVendaProduto: Erro=" & Err.Number & " - " & Err.Description)
    Call GravaAuditoria(1, Me.name, 25, "SubtraiVendaProduto: Erro inesperado...")
End Sub
Private Sub TelaAguarde(ByVal pMensagem As String, ByVal pMostraAguarde As Boolean)
    If pMostraAguarde Then
        frmAguarde.Show
        Call frmAguarde.MostraMensagens(pMensagem, Me.Caption, Me.Top, Me.Left, Me.Width, Me.Height)
        DoEvents
    Else
        Call frmAguarde.Finaliza
        Set frmAguarde = Nothing
    End If
End Sub
Function TestaCupomDemonstracao() As Boolean
    Dim dados As String
    
    lCupomDemonstracao = False
    dados = ReadINI("CUPOM FISCAL", "Cupom Demonstracao", gArquivoIni)
    If dados = "SIM" Then
        lCupomDemonstracao = True
    End If
    
    lEcfInstalada = False
    dados = ReadINI("CUPOM FISCAL", "ECF Instalada", gArquivoIni)
    If dados = "SIM" Then
        lEcfInstalada = True
    End If
    
    lImpBematech = False
    lImpSchalter = False
    lImpMecaf = False
    lImpQuick = False
    lImpDaruma = False
    lImpDarumaFW = False
   ' lImpDataregis = False '
    dados = ReadINI("CUPOM FISCAL", "Impressora Fiscal", gArquivoIni)
    If dados = "BEMATECH" Then
        lImpBematech = True
    ElseIf dados = ".BEMATECH" Then
        lImpBematech = True
    ElseIf dados = "SCHALTER" Then
        lImpSchalter = True
    ElseIf dados = "MECAF" Then
        lImpMecaf = True
    ElseIf dados = "QUICK" Then
        lImpQuick = True
    ElseIf dados = "DARUMA" Then
        lImpDaruma = True
    ElseIf dados = "DARUMAFW" Then
        lImpDarumaFW = True
    End If
    
    
    If lImpBematech Then
        'If EcfBematechReducaoZPendente Then
        '    MsgBox "Existe uma Reducao Z pendente." & Chr(10) & "O sistema irá imprimi-la automaticamente agora.", vbInformation, "Redução Z Pendente!"
        '    ImprimeReducaoZ
        'End If
    End If
    
    Exit Function
FileError:
    Exit Function
End Function
Function TestaConexao(ByVal pNumeroTeste As Integer, ByVal pVindoDe As String) As Boolean
    Call CriaLogCupom("TestaConexao - Iniciado. pNumeroTeste=" & pNumeroTeste & " - Vindo de: " & pVindoDe)
    TestaConexao = False
    If Aliquota.LocalizarUltimo() Then
        TestaConexao = True
        Call CriaLogCupom("TestaConexao - CONEXÃO OK. pNumeroTeste=" & pNumeroTeste)
    Else
        Call CriaLogCupom("TestaConexao - ERRO DE CONEXÃO. pNumeroTeste=" & pNumeroTeste)
        
        
        'If Conectar.ConexaoAtiva Then
        '    Call CriaLogCupom("TestaConexao - CONEXÃO ABERTA")
        'Else
        '    Call CriaLogCupom("TestaConexao - CONEXÃO FECHADA")
        'End If
        'Call CriaLogCupom("TestaConexao - DESTRUINDO Conectar")
        'Set Conectar = Nothing
        'Call CriaLogCupom("TestaConexao - DESTRUINDO ConectarNuvem")
        'Set ConectarNuvem = Nothing
        
        'Set Conectar = New CConexao
        'Set ConectarNuvem = New cConexaoNuvem
        
        
        'Teste pra tentar resolver problema que fecha conexao ao passar cartao pela seguna vez
        'Call CriaLogCupom("TestaConexao - Chama FinalizaConexao")
        'Conectar.FinalizaConexao
        'Call CriaLogCupom("TestaConexao - Concluida FinalizaConexao")
        '
        'Call CriaLogCupom("TestaConexao - Chama AbreConexao")
        'Conectar.AbreConexao
        'Call CriaLogCupom("TestaConexao - Concluida AbreConexao")
        '
        '
        'If Conectar.ConexaoAtiva Then
        '    Call CriaLogCupom("TestaConexao - CONEXÃO REABERTA")
        'Else
        '    Call CriaLogCupom("TestaConexao - CONEXÃO NÃO FOI REABERTA")
        'End If
    End If
    Call CriaLogCupom("TestaConexao - Finalizado. pNumeroTeste=" & pNumeroTeste)
End Function
Function TestaEmpresa() As Boolean

    'Dim dados As String
    Dim xNomeEmpresa As String
    'Dim NumeroArquivo As Integer
    Dim xAutomacao As Boolean
    Dim xTipoVenda As String
    
    On Error GoTo FileError
    
    'NumeroArquivo = FreeFile
    TestaEmpresa = False
    xAutomacao = False
    'Open "C:\VB5\SGP\CUPOM_DEMONSTRACAO.TXT" For Input As NumeroArquivo
    'If Not EOF(NumeroArquivo) Then
    '    Do Until EOF(NumeroArquivo)
    '        Line Input #NumeroArquivo, dados
    '        If Mid(dados, 1, 8) = "EMPRESA:" Then
    '            xNomeEmpresa = UCase(Mid(dados, 9, Len(dados) - 8))
    '        End If
    '        If Mid(dados, 1, 14) = "TIPO DE VENDA:" Then
    '            If UCase(Mid(dados, 15, Len(dados) - 8)) = "AUTOMACAO" Then
    '                xAutomacao = True
    '            End If
    '        End If
    '    Loop
    'End If
    'Close #NumeroArquivo
    
    
    xNomeEmpresa = ReadINI("CUPOM FISCAL", "Nome da Empresa", gArquivoIni)
    xTipoVenda = ReadINI("CUPOM FISCAL", "Tipo de Venda", gArquivoIni)
    If xTipoVenda = "AUTOMACAO" Or xTipoVenda = "AUTOMACAO/CONVENIENCIA" Then
        xAutomacao = True
    End If
    
    
    lTipoMovimento = 2
    If ReadINI("OUTRAS", "Apenas Visualizar", lNomeArquivoAutomacaoIni) = "NAO" Then
        If xAutomacao = False Then
            MsgBox "Este programa não pode ser executado neste computador!", vbInformation, "Erro de Configuração!"
            Exit Function
        End If
        If UCase(g_nome_empresa) = UCase(xNomeEmpresa) Then
            TestaEmpresa = True
        Else
            MsgBox "Este programa so pode ser executado quando a" & Chr(13) & "Empresa: " & xNomeEmpresa & Chr(13) & "Estiver selecionada!", vbInformation, "Erro de Consistencia!"
        End If
    Else
        TestaEmpresa = True
    End If
    Exit Function
FileError:
    Exit Function
End Function
Private Sub TestaEncerramentoCupomFiscal()
    Dim i As Integer
    Dim ACK As Integer
    Dim ST1 As Integer
    Dim ST2 As Integer
    Dim RetornoStatus As Integer
    Dim xString As String
    Dim xValor As String
    
    On Error GoTo FileError
    
    If lImpBematech Then
        i = 0
        'Call CriaLogCupom("Bematech_FI_FlagsFiscais(i)")
        'Bematech_FI_FlagsFiscais(i) 'NFCE - ALEX
        'Call CriaLogCupom("Bematech_FI_FlagsFiscais(i)=" & i & " - BemaRetorno=" & BemaRetorno)
        ' >= 33 = Ecf Aberto
        If i >= 33 And i < 128 Then
            TotalizaCupomAbertoNoBanco
            'Desconto para o Cupom Fiscal
            'Call CriaLogCupom("Bematech_FI_IniciaFechamentoCupom")
            'Bematech_FI_IniciaFechamentoCupom("D", "$", "00000000000000") ' NFCE - ALEX
            'Call CriaLogCupom("Bematech_FI_IniciaFechamentoCupom - BemaRetorno=" & BemaRetorno)
            'Verifica se o comando foi executado
            Call CriaLogCupom("Bematech_FI_RetornoImpressora(ACK, ST1, ST2)")
            RetornoStatus = 1 'Bematech_FI_RetornoImpressora(ACK, ST1, ST2) 'NFCE - ALEX
            Call CriaLogCupom("Bematech_FI_RetornoImpressora(ACK, ST1, ST2) ACM=" & ACK & " - ST1=" & ST1 & " - ST2=" & ST2)
             'Caso o comando não for executado inicia processo de fechamento do ecf
            If ST2 = 0 Then
                xValor = Space(14)
                'Busca SubTotal do Cupom Aberto
                'Call CriaLogCupom("Bematech_FI_SubTotal(xValor)=" & xValor)
                'Bematech_FI_SubTotal(xValor) 'NFCE - ALEX
                'Call CriaLogCupom("Bematech_FI_SubTotal - BemaRetorno=" & BemaRetorno)
                'Forma de Pagamento "Dinheiro"
                'Call CriaLogCupom("Bematech_FI_EfetuaFormaPagamento('Dinheiro        ', xValor=" & xValor)
                'Bematech_FI_EfetuaFormaPagamento("Dinheiro        ", xValor) 'NFCE - ALEX
                'Call CriaLogCupom("Bematech_FI_EfetuaFormaPagamento - BemaRetorno=" & BemaRetorno)
                'Fecha Cupom Fiscal
                'Call CriaLogCupom("Bematech_FI_TerminaFechamentoCupom('Cerrado Informatica - (062) 8436-4444           Sistemas para Automacao Comercial               ')")
                'Bematech_FI_TerminaFechamentoCupom("Cerrado Informatica - (062) 8436-4444           Sistemas para Automacao Comercial               ") 'NFCE-ALEX
                'Call CriaLogCupom("Bematech_FI_TerminaFechamentoCupom - BemaRetorno=" & BemaRetorno)
            End If
        End If
        xValor = 0
        xString = Space(6)
        'Call CriaLogCupom("Bematech_FI_NumeroCupom(xString)")
        'Bematech_FI_NumeroCupom(xString) 'NFCE - ALEX
        'Call CriaLogCupom("Bematech_FI_NumeroCupom xString=" & xString & " - BemaRetorno=" & BemaRetorno)
        'If BemaRetorno <> 1 Then
        '    Call AnalizaRetornoBematech(BemaRetorno)
        'End If
        lNumeroCupom = 1 'CLng(xString) + 1
'        With tbl_movimento_cupom_fiscal
'            .MoveLast
'            Do Until .BOF
'                If ![Numero do Cupom] <> lNumeroCupom Then
'                    Exit Do
'                End If
'                If ![Cupom Cancelado] = False And ![Item Cancelado] = False Then
'                    xValor = xValor + ![Valor Total]
'                End If
'                .MovePrevious
'            Loop
'
'            'Efetua Pagamento de Cupom Fiscal como Dinheiro
'            .MoveLast
'            Do Until .BOF
'                If ![Numero do Cupom] <> lNumeroCupom Then
'                    Exit Do
'                End If
'                .Edit
'                ![Forma de Pagamento] = 1
'                ![Valor Recebido] = xValor
'                .Update
'                .MovePrevious
'            Loop
'        End With
    End If
    Exit Sub

FileError:
    Call CriaLogCupom("Erro TestaEncerramentoCupomFiscal: Erro=" & Err.Number & " - " & Err.Description)
    MsgBox "Erro ao tentar fechar cupom fiscal em aberto", vbInformation, "TestaEncerramentoCupomFiscal"
    Exit Sub
End Sub
Private Sub TotalizaCupomAbertoNoBanco()
    Dim xNumeroCupom As Long
    Dim xData As Date
    Dim xOrdem As Integer
    Dim xValor As Currency
    
    xOrdem = 0
    xValor = 0
    If MovCupomFiscal.LocalizarUltimo(g_empresa, lNumeroPDV) Then
        xNumeroCupom = MovCupomFiscal.NumeroCupom
        xData = MovCupomFiscal.Data
    
        Do Until MovCupomFiscal.LocalizarNumeroProximaOrdem(g_empresa, lNumeroPDV, xNumeroCupom, xData, xOrdem) = False
            xValor = xValor + MovCupomFiscal.ValorTotal
            xOrdem = xOrdem + 1
        Loop
        If MovCupomFiscal.LocalizarCodigo(g_empresa, lNumeroPDV, xNumeroCupom, xData, 1) Then
            MovCupomFiscal.FormaPagamento = 1
            MovCupomFiscal.ValorRecebido = xValor
            If Not MovCupomFiscal.AlterarFormaPagamento(g_empresa, lNumeroPDV, xNumeroCupom, xData) Then
                MsgBox "Não foi possível alterar a forma de pagamento!", vbInformation, "Erro de Integridade"
            End If
        Else
            MsgBox "Não foi localizar o cupom fiscal", vbInformation, "Erro de Integridade"
        End If
    End If
End Sub
Private Function ImprimeCupomFiscal() As Boolean
    Dim xString As String
    Dim x_valor_desconto As Currency
    Dim x_valor_acrescimo As Currency
    Dim i As Integer
    Dim xACK As Integer
    Dim xST1 As Integer
    Dim xST2 As Integer
    
    Dim CodigoProduto As String
    Dim NomeProduto As String
    Dim xAliquota As String
    Dim Quantidade As String
    Dim Valor As String
    Dim ValorDesconto As String
    Dim ValorAcrescimo As String
    Dim Departamento As String
    Dim Un As String
    
    Dim xTruncaValor As Double
    Dim xTruncaQuantidade As Double
    Dim xTruncaTotalCalculado As Currency
    
    On Error GoTo FileError
    
    ImprimeCupomFiscal = False
    '25/06/14 ^
    If l_flag_cupom_fiscal = "F" Then
        l_flag_cupom_fiscal = "A"
        If lNotificacaoGic Then
            menu_personalizado.DesativaVerificacaoGIC
        End If
        cmdIniciaProcessoFinalizacaoNFCe.Enabled = True
        lCodigoFiscal = Aliquota.CodigoFiscal
        'mnuLeituraX.Enabled = False
        mnuPontoFuncionario.Enabled = False
        'Abre o cupom fiscal
        xString = ""
        If Val(l_codigo_cliente) > 0 Then
            If Val(Cliente.CGC) > 0 Then
                xString = Mid(Cliente.CGC, 1, 2) + "." + Mid(Cliente.CGC, 3, 3) + "." + Mid(Cliente.CGC, 6, 3) + "/" + Mid(Cliente.CGC, 9, 4) + "-" + Mid(Cliente.CGC, 13, 2)
            ElseIf Cliente.CPF <> "" Then
                xString = Mid(Cliente.CPF, 1, 3) + "." + Mid(Cliente.CPF, 4, 3) + "." + Mid(Cliente.CPF, 7, 3) + "-" + Mid(Cliente.CPF, 10, 2)
            End If
        End If
        If lImpBematech Then 'ALEX - NCFE
'                Call CriaLogCupom("Bematech_FI_AbreCupom(xString)")
'                BemaRetorno = Bematech_FI_AbreCupom(xString)
'                '25/06/14 variavel BemaRetoro tem que receber 1^
'                Call CriaLogCupom("Bematech_FI_AbreCupom(xString)=" & xString & " - BemaRetorno=" & BemaRetorno)
'                If BemaRetorno <> 1 Then
'                    Call CriaLogCupom("Bematech_FI_RetornoImpressora(xACK, xST1, xST2)")
'                    BemaRetorno = Bematech_FI_RetornoImpressora(xACK, xST1, xST2)
'                    Call CriaLogCupom("Bematech_FI_RetornoImpressora(xACK, xST1, xST2) xACM=" & xACK & " - xST1=" & xST1 & " - xST2=" & xST2)
'                    Call CriaLogCupom(Date & " " & Time & "???? ImprimeCupomFiscal: Bematech_FI_AbreCupom BemaRetorno=" & BemaRetorno & " - xACK=" & xACK & " - xST1=" & xST1 & " - xST2=" & xST2)
'                    l_flag_cupom_fiscal = "F"
'                    Exit Function
'                End If
'                Call CriaLogCupom("Bematech_FI_RetornoImpressora(xACK, xST1, xST2)")
'                BemaRetorno = Bematech_FI_RetornoImpressora(xACK, xST1, xST2)
'                Call CriaLogCupom("Bematech_FI_RetornoImpressora(xACK, xST1, xST2) xACM=" & xACK & " - xST1=" & xST1 & " - xST2=" & xST2)
'                If xST1 <> 0 Or xST2 <> 0 Then
'                    Call CriaLogCupom(Date & " " & Time & "???? ImprimeCupomFiscal: Bematech_FI_AbreCupom BemaRetorno=" & BemaRetorno & " - xACK=" & xACK & " - xST1=" & xST1 & " - xST2=" & xST2)
'                End If
            
        ElseIf lImpQuick Then 'ALEX - NCFE
'                If Val(l_codigo_cliente) > 0 Then
'                    xString = ""
'                    If Cliente.CGC <> "" Then
'                        xString = fMascaraCNPJ(Cliente.CGC)
'                    Else
'                        If Cliente.CPF <> "" Then
'                            xString = fMascaraCPF(Cliente.CPF)
'                        End If
'                    End If
'                    Call EcfQuickAbreCupomFiscal(dtcboCliente.Text, Cliente.Endereco & ", " & Cliente.Bairro & ", " & Cliente.Cidade, xString)
'                Else
'                    Call EcfQuickAbreCupomFiscal("", "", "")
'                End If
'                'Call EcfQuickVendeItem(True, -11, 0, "160", "", "Gasolina Comum", 0, 2.57, 10, "LT")
'                'Call EcfQuickCancelaCupom
        ElseIf lImpDaruma Then
'                'Abre Cupom Fiscal
'                xString = ""
'                If Val(l_codigo_cliente) > 0 Then
'                    If Cliente.CGC <> "" Then
'                        xString = fMascaraCNPJ(Cliente.CGC)
'                    Else
'                        If Cliente.CPF <> "" Then
'                            xString = fMascaraCPF(Cliente.CPF)
'                        End If
'                    End If
'                End If
'                BemaRetorno = Daruma_FI_AbreCupom(xString)
        ElseIf lImpDarumaFW Then
'                'Abre Cupom Fiscal
'                Dim xCpf As String
'                xString = ""
'                xCpf = ""
'                If Val(l_codigo_cliente) > 0 Then
'                    xString = Cliente.RazaoSocial
'                    If Cliente.CGC <> "" Then
'                        xCpf = fMascaraCNPJ(Cliente.CGC)
'                    Else
'                        If Cliente.CPF <> "" Then
'                            xCpf = fMascaraCPF(Cliente.CPF)
'                        End If
'                    End If
'                End If
'                BemaRetorno = iCFAbrir_ECF_Daruma(xCpf, xString, "")
'
        End If
    End If
    xString = Format(Format(fValidaValor(txt_valor_unitario.Text) * fValidaValor(txt_quantidade.Text), "###,##0.0000"), "###,##0.0000")
    i = Len(xString)
    xString = Mid(xString, 1, i - 2)
    x_valor_acrescimo = 0
    x_valor_desconto = 0
    If fValidaValor(txt_valor_total.Text) > fValidaValor(xString) Then
        x_valor_acrescimo = fValidaValor(txt_valor_total.Text) - fValidaValor(xString)
    ElseIf fValidaValor(txt_valor_total.Text) < fValidaValor(xString) Then
        x_valor_desconto = fValidaValor(xString) - fValidaValor(txt_valor_total.Text)
    Else
    End If
    'O teste abaixao é para evitar acrescimo errado em valores exagerado.
    'No Bosque aconteceu um acrescimo de 3.267,00 para um cupom que
    'deveria ser 33,00.
    If x_valor_acrescimo > 50 Then
        x_valor_acrescimo = 0
    End If
    If lImpBematech Then
        'código do produto
        CodigoProduto = Format(MovCupomFiscal.CodigoProduto, "#,##0")
        If Trim(Produto.CodigoBarra) <> "" Then
            CodigoProduto = Produto.CodigoBarra
        End If
        'nome do produto
        NomeProduto = Produto.Nome
        'tipo de tributação
        xAliquota = Aliquota.CodigoFiscal
        'Valor Unitário
        xString = Format(MovCupomFiscal.ValorUnitario, "000000.000")
        Valor = Mid(xString, 1, 6) + Mid(xString, 8, 3)
        'Quantidade
        xString = Format(MovCupomFiscal.Quantidade, "0000.000")
        Quantidade = Mid(xString, 1, 4) + Mid(xString, 6, 3)
        'Valor do Acréscimo
        xString = Format(x_valor_acrescimo, "00000000.00")
        ValorAcrescimo = Mid(xString, 1, 8) + Mid(xString, 10, 2)
        'Valor do Desconto
        xString = Format(x_valor_desconto, "00000000.00")
        ValorDesconto = Mid(xString, 1, 8) + Mid(xString, 10, 2)
        'Departamento
        If Aliquota.CodigoFiscal = "II" Then
            Departamento = Format(5, "00")
        ElseIf Aliquota.CodigoFiscal = "NN" Then
            Departamento = Format(6, "00")
        ElseIf Aliquota.CodigoFiscal = "FF" Then
            If Produto.CodigoGrupo = lGrupoCombustivel Then
                Departamento = Format(2, "00")
                'MsgBox "combustivel - 2"
            Else
                Departamento = Format(1, "00")
                'MsgBox "substituicao - 1"
            End If
        ElseIf Aliquota.Aliquota > 5 Then
            Departamento = Format(3, "00")
        ElseIf Aliquota.Aliquota > 0 And Aliquota.Aliquota <= 5 Then
            Departamento = Format(7, "00")
        End If
        'Unidade de Medida
        Un = Mid(Produto.Unidade, 1, 2)
        If lEcfTruncamento = True Then
            xTruncaValor = MovCupomFiscal.ValorUnitario
            If lEcfQtdCasasDecimais = 2 Then
                xTruncaQuantidade = Mid(Format(MovCupomFiscal.Quantidade, "0000000000.0000"), 1, 13)
            Else
                xTruncaQuantidade = MovCupomFiscal.Quantidade
            End If
            xTruncaTotalCalculado = fValidaValor(Mid(Format(xTruncaValor * xTruncaQuantidade, "0000000000.000000"), 1, 13))
            ValorAcrescimo = "0000000000"
            ValorDesconto = "0000000000"
            If fValidaValor(txt_valor_total.Text) > xTruncaTotalCalculado Then
                x_valor_acrescimo = fValidaValor(txt_valor_total.Text) - xTruncaTotalCalculado
                Call CriaLogCupom("Acrescimo Truncamento  txt_valor_total=" & txt_valor_total.Text & " xTruncaTotalCalculado=" & xTruncaTotalCalculado)
                xString = Format(x_valor_acrescimo, "00000000.00")
                ValorAcrescimo = Mid(xString, 1, 8) + Mid(xString, 10, 2)
            ElseIf fValidaValor(txt_valor_total.Text) < xTruncaTotalCalculado Then
                x_valor_desconto = xTruncaTotalCalculado - fValidaValor(txt_valor_total.Text)
                Call CriaLogCupom("Desconto Truncamento   txt_valor_total=" & txt_valor_total.Text & " xTruncaTotalCalculado=" & xTruncaTotalCalculado)
                xString = Format(x_valor_desconto, "00000000.00")
                ValorDesconto = Mid(xString, 1, 8) + Mid(xString, 10, 2)
            End If
        End If
        'Call CriaLogCupom("Bematech_FI_VendeItemDepartamento(CodigoProduto, NomeProduto... CodigoProduto=" & CodigoProduto & " - NomeProduto=" & NomeProduto)
        'ALEX - NFCE 'Bematech_FI_VendeItemDepartamento(CodigoProduto, NomeProduto, xAliquota, Valor, Quantidade, ValorAcrescimo, ValorDesconto, Departamento, Un)
        'Call CriaLogCupom("Bematech_FI_VendeItemDepartamento - BemaRetorno=" & BemaRetorno)
        'If BemaRetorno <> 1 Then
        '    Call AnalizaRetornoBematech(BemaRetorno)
        'End If
        'Call CriaLogCupom("Bematech_FI_RetornoImpressora(xACK, xST1, xST2)")
        'ALEX - NFCE'Bematech_FI_RetornoImpressora(xACK, xST1, xST2)
        'Call CriaLogCupom("Bematech_FI_RetornoImpressora(xACK, xST1, xST2) xACM=" & xACK & " - xST1=" & xST1 & " - xST2=" & xST2)
        'If BemaRetorno = 1 Then
            ImprimeCupomFiscal = True
        'Else
        '    Call CriaLogCupom(Date & " " & Time & "???? ImprimeCupomFiscal: Bematech_FI_VendeItemDepartamento BemaRetorno=" & BemaRetorno & " - xACK=" & xACK & " - xST1=" & xST1 & " - xST2=" & xST2)
        'End If
        'If xST1 <> 0 Or xST2 <> 0 Then
        '    Call CriaLogCupom(Date & " " & Time & "???? ImprimeCupomFiscal: Bematech_FI_VendeItemDepartamento BemaRetorno=" & BemaRetorno & " - xACK=" & xACK & " - xST1=" & xST1 & " - xST2=" & xST2)
        'End If
    ElseIf lImpQuick Then 'ALEX - NFCE
        'código do produto
        CodigoProduto = Format(MovCupomFiscal.CodigoProduto, "#,##0")
        'nome do produto
        NomeProduto = Produto.Nome
'            Call EcfQuickVendeItem(True, -2, 0, CodigoProduto, "", NomeProduto, 0, MovCupomFiscal.ValorUnitario, MovCupomFiscal.Quantidade, Produto.Unidade)
'
'            'Valor do Acréscimo/Desconto
'            If x_valor_acrescimo > 0 Or x_valor_desconto > 0 Then
'                Call EcfQuickAcresceItemFiscal(lOrdem, False, x_valor_acrescimo, x_valor_desconto)
'            End If
        ImprimeCupomFiscal = True
    ElseIf lImpDaruma Then 'ALEX - NFCE
        'código do produto
        CodigoProduto = Format(MovCupomFiscal.CodigoProduto, "#,##0")
        'nome do produto
        NomeProduto = Produto.Nome
        'tipo de tributação
        xAliquota = Aliquota.CodigoFiscal
        'Unidade de Medida
        Un = Mid(Produto.Unidade, 1, 2)
        'Quantidade
        Quantidade = Format(MovCupomFiscal.Quantidade, "0000.000")
        'Valor Unitário
        Valor = Format(MovCupomFiscal.ValorUnitario, "000000.000")
        'Valor do Acréscimo
        ValorAcrescimo = Format(x_valor_acrescimo, "00000000.00")
        'Valor do Desconto
        ValorDesconto = Format(x_valor_desconto, "00000000.00")
        If lEcfTruncamento = True Then
            xTruncaValor = MovCupomFiscal.ValorUnitario
            If lEcfQtdCasasDecimais = 2 Then
                xTruncaQuantidade = Mid(Format(MovCupomFiscal.Quantidade, "0000000000.0000"), 1, 13)
            Else
                xTruncaQuantidade = MovCupomFiscal.Quantidade
            End If
            xTruncaTotalCalculado = fValidaValor(Mid(Format(xTruncaValor * xTruncaQuantidade, "0000000000.000000"), 1, 13))
            ValorAcrescimo = "0000000000"
            ValorDesconto = "0000000000"
            If fValidaValor(txt_valor_total.Text) > xTruncaTotalCalculado Then
                x_valor_acrescimo = fValidaValor(txt_valor_total.Text) - xTruncaTotalCalculado
                Call CriaLogCupom("Acrescimo Truncamento  txt_valor_total=" & txt_valor_total.Text & " xTruncaTotalCalculado=" & xTruncaTotalCalculado)
                ValorDesconto = Format(x_valor_acrescimo * -1, "00000000.00")
            ElseIf fValidaValor(txt_valor_total.Text) < xTruncaTotalCalculado Then
                x_valor_desconto = xTruncaTotalCalculado - fValidaValor(txt_valor_total.Text)
                Call CriaLogCupom("Desconto Truncamento   txt_valor_total=" & txt_valor_total.Text & " xTruncaTotalCalculado=" & xTruncaTotalCalculado)
                ValorDesconto = Format(x_valor_desconto, "00000000.00")
            End If
        End If
        'Departamento
        Departamento = Format(1, "00")
        If Aliquota.CodigoFiscal = "II" Then
            Departamento = Format(5, "00")
        ElseIf Aliquota.CodigoFiscal = "NN" Then
            Departamento = Format(6, "00")
        ElseIf Aliquota.CodigoFiscal = "FF" Then
            If Produto.CodigoGrupo = lGrupoCombustivel Then
                Departamento = Format(2, "00")
                'MsgBox "combustivel - 2"
            Else
                Departamento = Format(1, "00")
                'MsgBox "substituicao - 1"
            End If
        ElseIf Aliquota.Aliquota > 5 Then
            Departamento = Format(3, "00")
        ElseIf Aliquota.Aliquota > 0 And Aliquota.Aliquota <= 5 Then
            Departamento = Format(7, "00")
        End If
        'ALEX - NFCE'Daruma_FI_VendeItem(CodigoProduto, NomeProduto, xAliquota, Un, Quantidade, 3, Valor, "$", ValorDesconto)
        'Venda por departamento, nao encontrei, parece que está em desuso
        'BemaRetorno = Daruma_FI_VendeItemDepartamento(CodigoProduto, NomeProduto, xAliquota, Valor, Quantidade, ValorAcrescimo, ValorDesconto, Departamento, Un)
        'If BemaRetorno = 1 Then
            ImprimeCupomFiscal = True
        'Else
        '    'DarumaBuscaRetorno 'ALEX - NFCE
        '    Call CriaLogCupom("???? ImprimeCupomFiscal: Daruma_FI_VendeItem BemaRetorno=" & BemaRetorno & " - lAck=" & lAck & " - lSt1=" & lSt1 & " - lSt2=" & lSt2 & " - lErroExtendido=" & lErroExtendido)
        'End If
    End If
'    Else
'        l_flag_cupom_fiscal = "A"
'        If lNotificacaoGic Then
'            menu_personalizado.DesativaVerificacaoGIC
'        End If
'        cmdIniciaProcessoFinalizacaoNFCe.Enabled = True
'        'mnuLeituraX.Enabled = False
'        mnuPontoFuncionario.Enabled = False
'        ImprimeCupomFiscal = True
'    End If
    Exit Function

FileError:
    Call CriaLogCupom("Erro ImprimeCupomFiscal: Erro=" & Err.Number & " - " & Err.Description)
    MsgBox "Não foi possível imprimir o novo cupom fiscal.", vbCritical, "Erro Grave!"
    Exit Function
End Function
Private Sub VerificaImprimeNotaAbastecimento()
    Dim xQtdVias As Integer
    Dim xQtdViasDocumentoVinculado As Integer
    
    If lImprimeDocumentoVinculado = False Or Val(cbo_forma_pagamento.Text) <> 5 Or l_codigo_cliente = 0 Then
        Exit Sub
    End If

    xQtdViasDocumentoVinculado = 1
    If lQtdViasConfDiv > 1 Then
        If g_nome_empresa = "POSTO PELICANO LTDA" Or g_nome_empresa = "POSTO PELICANO II LTDA" Or g_nome_empresa = "TEIXEIRA E PINHEIRO LTDA" Then
            If (MsgBox("Deseja imprimir a via do cliente?", vbQuestion + vbYesNo + vbDefaultButton2, "Imprime Via de Cliente?")) = vbYes Then
                xQtdViasDocumentoVinculado = lQtdViasConfDiv
            End If
        End If
    End If
    For xQtdVias = 1 To xQtdViasDocumentoVinculado
        ImpTermicaAbreRelatorio
        Call ImpTermicaCabecalhoPosto(lNumeroCupom)
        Call ImpTermicaDadosCliente
        Call ImpTermicaImprimeDados("------------------------------------------------", True)
        ImpTermicaFechaRelatorio ("Relatório de Notas de Abastecimento")
    Next
End Sub
''Private Sub ImprimeLeituraXCombustivel()
''    Dim xString As String
''    Dim xLinha As String
''    Dim i As Integer
''    Dim xSubSQL As String
''
''    On Error GoTo trata_erro
''
''    xString = "NOME DO COMBUSTIVEL      QUANTIDADE        VALOR"
''    lSQL = "SELECT Codigo, Nome FROM Combustivel WHERE Empresa = " & g_empresa & " ORDER BY Nome"
''    Call AtualizaRecordset(0)
''    With rst
''        If .RecordCount > 0 Then
''            .MoveFirst
''            Do Until .EOF
''                xSubSQL = ""
''                xSubSQL = xSubSQL & "SELECT [Codigo do Produto] "
''                xSubSQL = xSubSQL & "FROM Bomba "
''                xSubSQL = xSubSQL & "WHERE [Tipo de Combustivel] = " & Chr(39) & rst!Codigo & Chr(39) & " "
''                xSubSQL = xSubSQL & "GROUP BY [Codigo do Produto]"
''
''                lSQL = ""
''                lSQL = lSQL & "SELECT SUM(Quantidade) AS [Total Quantidade], SUM([Valor Total]) AS [Total Valor] "
''                lSQL = lSQL & "FROM Movimento_Cupom_Fiscal "
''                lSQL = lSQL & "WHERE [Data do Cupom] = #" & Format(Date, "mm/dd/yyyy") & "# "
''                lSQL = lSQL & "AND [Cupom Cancelado] = " & preparaBooleano(False)
''                lSQL = lSQL & "AND [Item Cancelado] = " & preparaBooleano(False)
''                lSQL = lSQL & "AND [Codigo do Produto] IN ( " & xSubSQL & " )"
''                Call AtualizaRecordset2(0)
''
''                         '         1         2         3         4       4
''                         '123456789012345678901234567890123456789012345678
''                'xLinha ="NOME DO COMBUSTIVEL      QUANTIDADE        VALOR"
''                'xLinha ="12345678901234567890   #,###,##0.00 #,###,##0.00"
''                xLinha = "A                              0,00         0,00"
''                Mid(xLinha, 1, 20) = rst!Nome
''                If Not rst2.EOF Then
''                    If Not IsNull(rst2("Total Quantidade").Value) And Not IsNull(rst2("Total Valor").Value) Then
''                        i = Len(Format(rst2("Total Quantidade").Value, "#,###,##0.00"))
''                        Mid(xLinha, 24 + 12 - i, i) = Format(rst2("Total Quantidade").Value, "#,###,##0.00")
''                        i = Len(Format(rst2("Total Valor").Value, "#,###,##0.00"))
''                        Mid(xLinha, 37 + 12 - i, i) = Format(rst2("Total Valor").Value, "#,###,##0.00")
''                    End If
''                End If
''
''
''                xString = xString & xLinha
''
''                rst2.Close
''                Set rst2 = Nothing
''                .MoveNext
''            Loop
''        End If
''        .Close
''    End With
''    Set rst = Nothing
''
''    'Abertura do Relatório Gerencial
''    Call CriaLogCupom("Bematech_FI_RelatorioGerencial(xString) - xString=" & xString)
''    BemaRetorno = Bematech_FI_RelatorioGerencial(xString)
''    Call CriaLogCupom("Bematech_FI_RelatorioGerencial - BemaRetorno=" & BemaRetorno)
''
''    'Fechamento de Relatório Gerencial
''    Call CriaLogCupom("Bematech_FI_FechaRelatorioGerencial")
''    BemaRetorno = Bematech_FI_FechaRelatorioGerencial
''    Call CriaLogCupom("Bematech_FI_FechaRelatorioGerencial - BemaRetorno=" & BemaRetorno)
''    Exit Sub
''
''trata_erro:
''    Call CriaLogCupom("Erro ImprimeLeituraXCombustivel: Erro=" & Err.Number & " - " & Err.Description)
''End Sub
'Private Sub ImprimeProgramaFormaPagamento()
''    Dim x_data  As Date
'    Dim i As Integer
'    Dim ACK As Integer
'    Dim ST1 As Integer
'    Dim ST2 As Integer
'    Dim x_string As String
'    Dim NumeroArquivo As Integer
'    Dim dados As String
'
'    On Error GoTo FileError
'
'    lExisteImpressora = True
''    If lImpQuick Then
''        If EcfQuickLeRegistrador("SemPapel", "Indicador", 0) = "0" Then
''            lExisteImpressora = True
''        Else
''            lExisteImpressora = False
''        End If
''    End If
'
'    'Comentei codigo abaixo pra melhorar performance ao abrir tela e não comunicar com dll da ECF - Tasso 20/04/2017
''    If lImpBematech Then
''        If Not Testa_ImpressoraCF Then
''            lExisteImpressora = False
''        End If
''    End If
'    'If lImpBematech And lExisteImpressora Then
''        If lTestaReducaoZpendente Then
''            lTestaReducaoZpendente = False
''            BematechReducaoZPendente
''        End If
''        If lNumeroPDV > 0 Then
''            GravaMapaResumo
''        End If
'
'        'Programa Nomeação de Departamento para o Cupom Fiscal
'        'Call CriaLogCupom("Bematech_FI_FlagsFiscais(i) - i=" & i)
'        'BemaRetorno = 1 'Bematech_FI_FlagsFiscais(i)
'        'Call CriaLogCupom("Bematech_FI_FlagsFiscais(i)=" & i & " - BemaRetorno=" & BemaRetorno)
'        'If i <> 1 And i <> 5 And i <> 37 Then
'            'If i = 32 Then    NAO É O CODIGO 32
'            '    MsgBox "APARENTEMENTE existe uma pendência de Redução Z.", vbInformation, "Redução Z Pendente!"
'            '    If (MsgBox("Deseja imprimir a Redução Z pendente?", vbQuestion + vbYesNo + vbDefaultButton2, "Imprime Redução Z Pendente?")) = vbYes Then
'            '        ImprimeReducaoZ
'            '    End If
'            'End If
''            Call CriaLogCupom("Bematech_FI_NomeiaDepartamento(1, 'RETIDOS   ')")
''            BemaRetorno = Bematech_FI_NomeiaDepartamento(1, "RETIDOS   ")
''            Call CriaLogCupom("Bematech_FI_NomeiaDepartamento - BemaRetorno=" & BemaRetorno)
'
''            Call CriaLogCupom("Bematech_FI_RetornoImpressora(ACK, ST1, ST2)")
''            BemaRetorno = Bematech_FI_RetornoImpressora(ACK, ST1, ST2)
''            Call CriaLogCupom("Bematech_FI_RetornoImpressora(ACK, ST1, ST2) ACM=" & ACK & " - ST1=" & ST1 & " - ST2=" & ST2 & " - BemaRetorno=" & BemaRetorno)
''            If ST1 = 0 And ST2 = 0 Then
''                Call CriaLogCupom("Bematech_FI_NomeiaDepartamento(2, 'COMBUST.  ')")
''                BemaRetorno = Bematech_FI_NomeiaDepartamento(2, "COMBUST.  ")
''                Call CriaLogCupom("Bematech_FI_NomeiaDepartamento - BemaRetorno=" & BemaRetorno)
''                Call CriaLogCupom("Bematech_FI_NomeiaDepartamento(3, 'TRIBUTADO ')")
''                BemaRetorno = Bematech_FI_NomeiaDepartamento(3, "TRIBUTADO ")
''                Call CriaLogCupom("Bematech_FI_NomeiaDepartamento - BemaRetorno=" & BemaRetorno)
''                Call CriaLogCupom("Bematech_FI_NomeiaDepartamento(4, 'AFERICAO  ')")
''                BemaRetorno = Bematech_FI_NomeiaDepartamento(4, "AFERICAO  ")
''                Call CriaLogCupom("Bematech_FI_NomeiaDepartamento - BemaRetorno=" & BemaRetorno)
''                Call CriaLogCupom("Bematech_FI_NomeiaDepartamento(5, 'ISENTO    ')")
''                BemaRetorno = Bematech_FI_NomeiaDepartamento(5, "ISENTO    ")
''                Call CriaLogCupom("Bematech_FI_NomeiaDepartamento - BemaRetorno=" & BemaRetorno)
''                Call CriaLogCupom("Bematech_FI_NomeiaDepartamento(6, 'NAO INCID.')")
''                BemaRetorno = Bematech_FI_NomeiaDepartamento(6, "NAO INCID.")
''                Call CriaLogCupom("Bematech_FI_NomeiaDepartamento - BemaRetorno=" & BemaRetorno)
''                Call CriaLogCupom("Bematech_FI_NomeiaDepartamento(7, 'SERVICOS  ')")
''                BemaRetorno = Bematech_FI_NomeiaDepartamento(7, "SERVICOS  ")
''                Call CriaLogCupom("Bematech_FI_NomeiaDepartamento - BemaRetorno=" & BemaRetorno)
''            End If
'        'End If
''        x_string = Space(1)
''        Call CriaLogCupom("Bematech_FI_VerificaTruncamento(x_string) - x_string" & x_string)
''        BemaRetorno = Bematech_FI_VerificaTruncamento(x_string)
''        Call CriaLogCupom("Bematech_FI_VerificaTruncamento" & " - x_string" & x_string & " - BemaRetorno=" & BemaRetorno)
'        'lEcfTruncamento = False
''        If x_string = "1" Then
''            lEcfTruncamento = True
''        End If
'    'End If
''    If lImpQuick And lExisteImpressora Then
''        If Not EcfQuickLeMeioPagamento("Nota Vinculada") Then
''            If Not EcfQuickDefineMeioPagamento("Nota Vinculada", "Nota de Abastecimento com opcao de ser vinculada ao ecf.", True) Then
''                MsgBox "Erro ao programar forma de pagamento (Nota Vinculada).", vbCritical, "Erro Inesperado!"
''            End If
''        End If
''        If Not EcfQuickLeMeioPagamento("TEF") Then
''            If Not EcfQuickDefineMeioPagamento("TEF", "Transferencia Eletronica de Fundos.", True) Then
''                MsgBox "Erro ao programar forma de pagamento (TEF).", vbCritical, "Erro Inesperado!"
''            End If
''        End If
'''        If Not EcfQuickLeMeioPagamento("Cheque TEF") Then
'''            If Not EcfQuickDefineMeioPagamento("Cheque TEF", "Cheque - TEF", True) Then
'''                MsgBox "Erro ao programar forma de pagamento (Cheque TEF).", vbCritical, "Erro Inesperado!"
'''            End If
'''        End If
'''        If Not EcfQuickLeMeioPagamento("C H E Q U E") Then
'''            If Not EcfQuickDefineMeioPagamento("C H E Q U E", "C H E Q U E", True) Then
'''                MsgBox "Erro ao programar forma de pagamento (C H E Q U E TEF).", vbCritical, "Erro Inesperado!"
'''            End If
'''        End If
''        If Not EcfQuickLeMeioPagamento("CONSULTA CHEQUE") Then
''            If Not EcfQuickDefineMeioPagamento("CONSULTA CHEQUE", "CONSULTA CHEQUE", True) Then
''                MsgBox "Erro ao programar forma de pagamento (CONSULTA CHEQUE TEF).", vbCritical, "Erro Inesperado!"
''            End If
''        End If
''
''        GravaMapaResumo
''    End If
''    If lImpDaruma And lExisteImpressora Then
'''        BemaRetorno = Daruma_FI_VerificaImpressoraLigada
'''
'''        'Verifica e emite redução Z pendente
'''        dados = Space(2)
'''        BemaRetorno = Daruma_FI_VerificaZPendente(dados)
'''        If Mid(dados, 1, 1) = "1" Then
'''            MsgBox "Existe uma redução Z pendente.", vbInformation, "Redução Z Pendente."
'''            'ImprimeReducaoZ
'''        End If
''
''        'Verifica se reducao Z foi impressa no dia
''        'Aqui não funciona, pois se a redução Z pendente sair
''        'no início do dia, a xDataRdz vem com data do dia atual.
'''        Dim xDataRdz As String
'''        Dim xHoraRdz As String
'''        xDataRdz = Space(6)
'''        xHoraRdz = Space(6)
'''        BemaRetorno = Daruma_FI_DataHoraReducao(xDataRdz, xHoraRdz)
'''        xDataRdz = Mid(xDataRdz, 1, 2) & "/" & Mid(xDataRdz, 3, 2) & "/20" & Mid(xDataRdz, 5, 2)
'''        If CDate(xDataRdz) = Date Then
'''            MsgBox "Reducao Z já foi impressa." & Chr(10) & "Será aceito imprimir Cupom Fiscal somente após: " & Format(Date, "dd/mm/yyyy") & Chr(10) & "O Sistema Será Fechado Automaticamente.", vbInformation, "Fechando o Sistema!"
'''            If g_nivel_acesso > 1 Then
'''                End
'''            End If
'''        End If
''
''        '28/04
''        GravaMapaResumo
''
''        'dados = Space(2)
''        'BemaRetorno = Daruma_FI_VerificaTruncamento(dados)
''        lEcfTruncamento = False
''        'MsgBox "Truncamento = >" & dados & "<"
''        'If Mid(dados, 1, 1) = "1" Then
''        '    lEcfTruncamento = True
''        'End If
''
''        'BemaRetorno = Daruma_FI_ProgramaFormasPagamento("Ch. A Vista;Ch. Pre-Datado;Cartao Credito;Nota Vinculada;TEF")
''    End If
''    If lImpDarumaFW And lExisteImpressora Then
'''        BemaRetorno = eDefinirProduto_Daruma("ECF")
'''        BemaRetorno = regAlterarValor_Daruma("ECF\RetornarAvisoErro", "1")
'''        BemaRetorno = eBuscarPortaVelocidade_ECF_Daruma()
'''        If BemaRetorno = 1 Then
'''            MsgBox "Problema ao comunicar dom Daruma FW", vbOKOnly + vbCritical, "Erro de Comunicação com ECF!"
'''        End If
''
''        'Verifica e emite redução Z pendente
'''        Dim xZPendente As String
'''        xZPendente = Space(1)
'''        BemaRetorno = rVerificarReducaoZ_ECF_Daruma(xZPendente)
'''        If xZPendente = "1" Then
'''            MsgBox "Existe uma redução Z pendente.", vbInformation, "Redução Z Pendente."
'''            'ImprimeReducaoZ
'''        End If
''
''
''        ' AQUI ESTUDAR COMO SABER SE ECF ESTÁ NO MODO TRUNCAMENTO
''        'dados = Space(2)
''        'BemaRetorno = Daruma_FI_VerificaTruncamento(dados)
''        lEcfTruncamento = False
''        'If Mid(dados, 1, 1) = "1" Then
''        '    lEcfTruncamento = True
''        'End If
''
''        ' AQUI ESTUDAR COMO PROGRAMAR MEIOS DE PAGAMENTO
''        'BemaRetorno = Daruma_FI_ProgramaFormasPagamento("Ch. A Vista;Ch. Pre-Datado;Cartao Credito;Nota Vinculada;TEF")
''    End If
'
'    Exit Sub
'
'FileError:
'    Call CriaLogCupom("Erro ImprimeProgramaFormaPagamento: Erro=" & Err.Number & " - " & Err.Description)
''    If lCupomDemonstracao = False Then
''        MsgBox "Não foi possível programar as formas de pagamento para o cupom fiscal.", vbCritical, "Erro Grave!"
''        Finaliza
''    End If
'End Sub
'Private Sub ImprimeReducaoZ()
'    Dim xRetorno As Long
'    Dim xData As String
'    Dim xHora As String
'
'    If lImpBematech Then
'        xData = Format(Date, "dd/mm/yyyy")
'        xHora = Format(Time, "hh:mm:ss")
'        Call CriaLogCupom("Bematech_FI_ReducaoZ(xData, xHora) - xData=" & xData & " - xHora=" & xHora)
'        BemaRetorno = 1 'Bematech_FI_ReducaoZ(xData, xHora) 'NFCE - ALEX
'        Call CriaLogCupom("Bematech_FI_ReducaoZ - BemaRetorno=" & BemaRetorno)
'    ElseIf lImpDaruma Then
'        xData = Format(Date, "dd/mm/yyyy")
'        xHora = Format(Time, "hh:mm:ss")
'        BemaRetorno = Daruma_FI_ReducaoZAjustaDataHora(xData, xHora)
'    '01/04/2016
'    ElseIf lImpQuick Then
'        BemaRetorno = EcfQuickReducaoZ
'    End If
'End Sub
Private Sub ImprimeResumoVendas()
    Dim x_string As String
    Dim x_sub_quantidade As Currency
    Dim x_sub_total As Currency
    Dim x_total_quantidade As Currency
    Dim x_total As Currency
    Dim x_linha As Integer
    Dim i As Integer
    Dim xSQL As String
    Dim rsProduto As New ADODB.Recordset
    Dim rsMovCupomFiscal As New ADODB.Recordset
    
    On Error GoTo FileError
    x_linha = 0
    x_total_quantidade = 0
    x_total = 0
    
    
    'Prepara SQL
    xSQL = ""
    xSQL = xSQL & "   SELECT Produto.Codigo, Produto.Nome"
    xSQL = xSQL & "     FROM Produto"
    xSQL = xSQL & " ORDER BY Nome ASC"
    'Abre RecordSet
    Set rsProduto = New ADODB.Recordset
    Set rsProduto = Conectar.RsConexao(xSQL)
    
    
    If rsProduto.RecordCount > 0 Then
        rsProduto.MoveFirst
        Do Until rsProduto.EOF
            'Prepara SQL
            xSQL = ""
            xSQL = xSQL & "   SELECT SUM(Movimento_Cupom_Fiscal.Quantidade) AS TotalQuantidade, SUM(Movimento_Cupom_Fiscal.[Valor Total]) AS TotalValor"
            xSQL = xSQL & "     FROM Movimento_Cupom_Fiscal"
            xSQL = xSQL & "    WHERE [Codigo do Produto] = " & rsProduto("Codigo").Value
            xSQL = xSQL & "      AND Data = " & preparaData(CDate(g_data_def))
            'Abre RecordSet
            Set rsMovCupomFiscal = New ADODB.Recordset
            Set rsMovCupomFiscal = Conectar.RsConexao(xSQL)
            If rsMovCupomFiscal.RecordCount > 0 Then
                rsMovCupomFiscal.MoveFirst
                If Not IsNull(rsMovCupomFiscal("TotalValor").Value) Then
                    x_total_quantidade = x_total_quantidade + rsMovCupomFiscal("TotalQuantidade").Value
                    x_total = x_total + rsMovCupomFiscal("TotalValor").Value
                    If x_linha = 0 Then
                        Open "A:\RESUMO_VENDA_" & Format(Day(g_data_def), "00") & "_" & Format(Month(g_data_def), "00") & "_" & Year(g_data_def) & ".TXT" For Output As #3
                        x_string = "12345678901234567890123456789012345678901234567890123456789012345678901234567890"
                        x_string = "         1         2         3         4         5         6         7         8"
                        x_string = "+------------------------------------------------------------------------------+"
                        Print #3, x_string
                        x_string = "|                                                                 , __/__/____ |"
                        Mid(x_string, 3, 40) = g_nome_empresa
                        i = Len(g_cidade_empresa)
                        Mid(x_string, 37 + 30 - i, i) = g_cidade_empresa
                        Mid(x_string, 69, 10) = Format(lData, "dd/mm/yyyy")
                        Print #3, x_string
                        x_string = "| RESUMO DAS VENDAS                                           HORA,   __:__:__ |"
                        Mid(x_string, 71, 8) = Time
                        Print #3, x_string
                        x_string = "+----------------------------------------+------------+-----------+------------+"
                        Print #3, x_string
                        x_string = "|PRODUTO                                 |VL. UNITARIO|   QUANT.  |VALOR  TOTAL|"
                        Print #3, x_string
                        x_string = "+----------------------------------------+------------+-----------+------------+"
                        Print #3, x_string
                    End If
                    x_linha = x_linha + 1
                    x_string = "|                                        |            |           |            |"
                    Mid(x_string, 2, 40) = rsProduto("Nome").Value
                    i = Len(Format(rsProduto("Preco de Venda").Value, "####,##0.00"))
                    Mid(x_string, 43 + 11 - i, i) = Format(rsProduto("Preco de Venda").Value, "####,##0.00")
                    i = Len(Format(x_sub_quantidade, "###,##0.00"))
                    Mid(x_string, 56 + 10 - i, i) = Format(x_sub_quantidade, "###,##0.00")
                    i = Len(Format(x_sub_total, "####,##0.00"))
                    Mid(x_string, 68 + 11 - i, i) = Format(x_sub_total, "####,##0.00")
                    Print #3, x_string
                End If
            End If
            rsProduto.MoveNext
        Loop
    End If
    If x_linha > 0 Then
        x_string = "+----------------------------------------+------------+-----------+------------+"
        Print #3, x_string
        x_string = "|                            *** TOTAL   |            |           |            |"
        i = Len(Format(rsMovCupomFiscal("TotalQuantidade").Value, "###,##0.00"))
        Mid(x_string, 56 + 10 - i, i) = Format(rsMovCupomFiscal("TotalQuantidade").Value, "###,##0.00")
        i = Len(Format(rsMovCupomFiscal("TotalValor").Value, "####,##0.00"))
        Mid(x_string, 68 + 11 - i, i) = Format(rsMovCupomFiscal("TotalValor").Value, "####,##0.00")
        Print #3, x_string
        x_string = "+--- Cerrado Informatica. ---------------+------------+-----------+------------+"
        Print #3, x_string
        Close #3
    End If
    Set rsMovCupomFiscal = Nothing
    Set rsProduto = Nothing
    Exit Sub
    
FileError:
    Call CriaLogCupom("Erro ImprimeResumoVendas: Erro=" & Err.Number & " - " & Err.Description)
    MsgBox "Não foi possível imprimir o novo cupom fiscal.", vbCritical, "Erro Grave!"
    Exit Sub
End Sub
Private Sub AbilitaMenu(ByVal pAbilita As Boolean)
    mnuCaixaPista.Enabled = pAbilita
    mnuConsulta.Enabled = pAbilita
    mnuFuncaoADM.Enabled = pAbilita
    'mnuLeituraX.Enabled = pAbilita
    mnuFuncao.Enabled = pAbilita
    mnuSenha.Enabled = pAbilita
    mnuNFCe.Enabled = pAbilita
    If lMarcaAutomacao = "." Then
        mnuFechamentoCaixa.Enabled = False
        mnuLancamentoEncerrante.Enabled = False
    Else
        mnuFechamentoCaixa.Enabled = pAbilita
        mnuLancamentoEncerrante.Enabled = pAbilita
    End If
    mnRelatorio.Enabled = pAbilita
        
    If pAbilita = True Then
        If lCaixaIndividual = False Then
           mnRelatorio.Enabled = False
        Else
            If ConfiguracaoDiversa.LocalizarCodigo(1, "NFCEAUTO: Bloqueia Rel. Abast. Func.") Then
                mnAbastecimentoAutomacaoFuncionario.Enabled = Not ConfiguracaoDiversa.Verdadeiro
            End If
            If ConfiguracaoDiversa.LocalizarCodigo(1, "NFCEAUTO: Bloqueia Rel. Falta de Caixa") Then
                mnFaltaCaixa.Enabled = Not ConfiguracaoDiversa.Verdadeiro
            End If
        End If
    End If

End Sub
Private Sub AbortarNFCeAtual()
    If l_flag_cupom_fiscal = "A" Then
        If (MsgBox("Deseja realmente abortar a NFCe?", vbQuestion + vbYesNo + vbDefaultButton2, "Abortar NFCe")) = vbYes Then
        
            If MovDocEletronicoCabecalho Is Nothing Then Exit Sub
            
            If MovDocEletronicoCabecalho.CodigoUltimoEvento >= Val(EVENTO_NFCE.FECHADA) Then
                Call MsgBox("O documento não pode ser abandonado, pois já está sendo processado", vbInformation + vbOKOnly + vbDefaultButton1, "Operação negada")
            End If
        
            Call GravaAuditoria(1, Me.name, 25, "Abortar NFCe:" & lNumeroCupom) ' MovDocEletronicoCabecalho.numero
        
            If GravaDocumentoEletronicoEvento(MovDocEletronicoCabecalho, EVENTO_NFCE.ABANDONADA) Then
                Call LoopEstornaVendaProdutoNFCeAtual
                If RemoveItensNFCeAbandonada(MovDocEletronicoCabecalho) Then
                    Call MovDocEletronicoCabecalho.Excluir(MovDocEletronicoCabecalho.IdEstabelecimento, MovDocEletronicoCabecalho.DataEmissao, MovDocEletronicoCabecalho.Entrada, MovDocEletronicoCabecalho.Saida, MovDocEletronicoCabecalho.Modelo, MovDocEletronicoCabecalho.Serie, lNumeroCupom)
                    If g_automacao Then
                        Call LoopRetornaAbastecimentoAutomacao
                    End If
                    
                    l_flag_cupom_fiscal = "F"
                    
                    Call MovCupomFiscal.ExcluirNumeroData(MovDocEletronicoCabecalho.IdEstabelecimento, lNumeroCupom, MovDocEletronicoCabecalho.DataEmissao)
                    Call MovCupomFiscalItem.ExcluirNumeroData(MovDocEletronicoCabecalho.IdEstabelecimento, lNumeroCupom, MovDocEletronicoCabecalho.DataEmissao)
                    
                    lInformaFormaPagamento = False
                    
                    Call CancelarFinalizacaoNFCe
                    
                    If lNotificacaoGic Then
                        menu_personalizado.AtivaVerificacaoGIC
                    End If
        
                    lCodigoFiscal = "  "
                    
                    lNumeroCupom = 0
                    lNumeroUltimoCupom = 0
                    lTotalCupom = 0
                    lDescontoEspecial = 0
                    lOrdem = 0
                    
                    Call BuscaDados
                    
                    MostraTelaSenha
                End If
            Else
                MsgBox "Não foi possível abortar NFCe!", vbCritical, "Erro de Integridade"
            End If
        End If
     End If
End Sub


Private Sub AdicionaEstoque(ByVal pCodigoProduto As Long, ByVal pQuantidade As Currency, ByVal pTipoSubEstoque As Integer)
On Error GoTo trata_erro
    
    If Estoque.LocalizarCodigo(g_empresa, pCodigoProduto) Then
        Estoque.Quantidade = Estoque.Quantidade + pQuantidade
        If Estoque.Alterar(g_empresa, pCodigoProduto) Then
            If SubEstoque.AlterarQuantidade(g_empresa, pCodigoProduto, pTipoSubEstoque, pQuantidade, True) Then
            Else
                Call CriaLogCupom("Erro AdicionaEstoque:Sub-Estoque não alterado. Produto=" & pCodigoProduto & " Quantidade=" & pQuantidade & " SubEst=" & pTipoSubEstoque)
                MsgBox "Não foi possível alterar o sub-estoque!", vbInformation, "Erro de Integridade!"
            End If
        Else
            Call CriaLogCupom("Erro AdicionaEstoque:Estoque não alterado. Produto=" & pCodigoProduto & " Quantidade=" & pQuantidade & " SubEst=" & pTipoSubEstoque)
            MsgBox "Não foi possível alterar o estoque!", vbInformation, "Erro de Integridade!"
        End If
    Else
        Call CriaLogCupom("Erro AdicionaEstoque:Estoque não cadastrado. Produto=" & pCodigoProduto & " Quantidade=" & pQuantidade & " SubEst=" & pTipoSubEstoque)
        MsgBox "Estoque não cadastrado.", vbInformation, "Erro de Verificação!"
    End If
    Exit Sub

trata_erro:
    Call CriaLogCupom("Erro AdicionaEstoque:Desconhecido. Produto=" & pCodigoProduto & " Quantidade=" & pQuantidade & " SubEst=" & pTipoSubEstoque)
    Call CriaLogCupom("Erro AdicionaEstoque: Erro=" & Err.Number & " - " & Err.Description)
    MsgBox Err.Number & " - " & Err.Description
End Sub
Private Sub AlteraMovBombaParaSubCaixa(ByVal pData As Date, ByVal pPeriodo As Integer)
    Dim xNovoSubCaixa As Integer
    
    If MovimentoBomba.ExisteDataPeriodoSubCx(g_empresa, pData, pPeriodo, 999) Then
        xNovoSubCaixa = MovimentoBomba.ProximoSubCaixa(g_empresa, pData, pPeriodo)
        If MovimentoBomba.AlteraParaSubCaixa(g_empresa, pData, pPeriodo, xNovoSubCaixa) Then
            If Not MovCaixaPista.AlteraSubCaixaMovBomba(g_empresa, pData, pPeriodo, xNovoSubCaixa) Then
                MsgBox "Erro ao mudar mov.CaixaPista. para SubCaixa=" & xNovoSubCaixa, vbCritical, "Erro de Integridade!"
            End If
            If Not MovimentoBombaEscritorio.AlteraParaSubCaixa(g_empresa, pData, pPeriodo, xNovoSubCaixa) Then
                MsgBox "Erro ao mudar mov.BombaEscrit. para SubCaixa=" & xNovoSubCaixa, vbCritical, "Erro de Integridade!"
            End If
        Else
            MsgBox "Erro ao mudar mov.BombaCupom para SubCaixa=" & xNovoSubCaixa, vbCritical, "Erro de Integridade!"
        End If
    End If
End Sub
'
'Private Sub AtivaReducaoZ()
'    Dim xNivelAcesso As Integer
'
'    If ConfiguracaoDiversa.LocalizarCodigo(1, "ECF: Liberar Reducao Z Nivel") Then
'        xNivelAcesso = ConfiguracaoDiversa.Codigo
'    End If
'    If g_nivel_acesso > xNivelAcesso Then
'        mnuReducaoZ.Visible = False
'    Else
'        mnuReducaoZ.Visible = True
'    End If
'End Sub

Private Sub AtivaDesativaAguarde(ByVal pMensagem As String, ByVal pMostraAguarde As Boolean)
    Call CriaLogCupom("AtivaDesativaAguarde - pMensagem->" & pMensagem & "<-")
    frameAguarde.Top = 0
    frameAguarde.Left = 0
    frameAguarde.Height = 500
    frameAguarde.Width = 500
    frameAguarde.Enabled = pMostraAguarde
    If pMostraAguarde Then
        frameAguarde.Visible = True
        frameAguarde.ZOrder 0
        
        'Call MostraMensagensAguarde(pMensagem, Me.Caption, 10, 50, 13450, 8150)
        Call MostraMensagensAguarde(pMensagem, Me.Caption, 10, 30, 13755, 9955)
        DoEvents
    Else
        frameAguarde.Visible = False
        frameAguarde.ZOrder 1
        TimerAguarde.Interval = 0
        TimerAguarde.Enabled = False
    End If
End Sub
Private Sub AtivaDesativaBicos(ByVal pAtiva As Boolean)
    Dim i As Integer
    For i = 0 To lQtdBomba - 1
        cmd_bico(i).Enabled = pAtiva
    Next
End Sub
Private Sub AtivaDesativaTimer(ByVal xAtiva As Boolean)
    If xAtiva Then
        Timer2.Enabled = True
        Timer2.Interval = 30
    Else
        Timer2.Enabled = False
        Timer2.Interval = 0
    End If
End Sub
Private Function AguardaSolicitacaoAutomacao() As Boolean
    Dim xHoraInicial As Date
    Dim i As Integer

    AguardaSolicitacaoAutomacao = False
    lbl_mensagem.Caption = "Aguarde... Verificando Serviço."
    
    xHoraInicial = Time
    'Fica até 7 segundos
    Do Until DateDiff("s", xHoraInicial, Time) >= 7
        If SolicitacaoFuncaoAutomacao.VerificaSeEstaEmAnalise(lNSU) Then
            AguardaSolicitacaoAutomacao = True
            lbl_mensagem.Caption = SolicitacaoFuncaoAutomacao.BuscaMensagem(lNSU)
            DoEvents
            Exit Do
        End If
        lbl_mensagem.Caption = SolicitacaoFuncaoAutomacao.BuscaMensagem(lNSU)
        DoEvents
        Call AguardaMS(500)
    Loop
    If AguardaSolicitacaoAutomacao = False Then
        If SolicitacaoFuncaoAutomacao.DefineHoraCancelamentoAC(lNSU, Time) Then
            MsgBox "Tempo de solicitação de Serviço de Automação excedido.", vbCritical, "Tempo Excedido!"
        Else
            MsgBox "Não será possível definir cancelamento de Solicitação pela AC.", vbCritical, "Erro de Integridade!"
        End If
    End If
End Function
Private Function AguardaSolicitacaoPetromovelAuto(ByVal pTempoEspera As Integer, ByVal pNSU As Long) As Boolean
    Dim xHoraInicial As Date
    Dim i As Integer
    Dim xMovSolicitacao As New CadastroDLL.cMovSolicitacaoFuncaoNFe

    AguardaSolicitacaoPetromovelAuto = False
    lbl_mensagem.Caption = "Aguarde... Verificando Serviço PetromovelAuto."
    DoEvents
    
    xHoraInicial = Time
    'Fica o tempo definido por parâmetro
    Do Until DateDiff("s", xHoraInicial, Time) >= pTempoEspera
        If xMovSolicitacao.VerificaSeEstaEmAnalise(pNSU) Then
            AguardaSolicitacaoPetromovelAuto = True
            lbl_mensagem.Caption = xMovSolicitacao.BuscaMensagem(pNSU)
            DoEvents
            Exit Do
        End If
        lbl_mensagem.Caption = xMovSolicitacao.BuscaMensagem(pNSU)
        DoEvents
        Call AguardaMS(500)
    Loop
    If AguardaSolicitacaoPetromovelAuto = False Then
        If xMovSolicitacao.DefineHoraCancelamentoAC(pNSU, Time) Then
            'MsgBox "Tempo de solicitação de Serviço do PetromovelAuto excedido.", vbCritical, "Tempo Excedido!"
            Call CriaLogSGP("[AguardaSolicitacaoPetromovelAuto]", "Tempo de solicitação de Serviço do PetromovelAuto excedido.", "")
            lbl_mensagem.Caption = "Tempo de solicitação de Serviço do PetromovelAuto excedido."
            DoEvents
        Else
            'MsgBox "Não será possível definir cancelamento de Solicitação pela AC.", vbCritical, "Erro de Integridade!"
            Call CriaLogSGP("[AguardaSolicitacaoPetromovelAuto]", "Não será possível definir cancelamento de Solicitação pela AC.", "")
            lbl_mensagem.Caption = "Não será possível definir cancelamento de Solicitação pela AC."
            DoEvents
        End If
    End If
End Function
Private Function AguardaSolicitAutoAprovado(ByVal pSegundos As Integer) As Boolean
    Dim xHoraInicial As Date
    Dim i As Integer

    AguardaSolicitAutoAprovado = False
    lbl_mensagem.Caption = "Aguarde... Verificando Serviço."
    
    xHoraInicial = Time
    'Fica até pSegundos
    Do Until DateDiff("s", xHoraInicial, Time) >= pSegundos
        If SolicitacaoFuncaoAutomacao.VerificaSeEstaAprovado(lNSU) Then
            AguardaSolicitAutoAprovado = True
            lbl_mensagem.Caption = SolicitacaoFuncaoAutomacao.BuscaMensagem(lNSU)
            DoEvents
            Exit Do
        ElseIf SolicitacaoFuncaoAutomacao.VerificaSeEstaCanceladoHost(lNSU) Then
            AguardaSolicitAutoAprovado = False
            lbl_mensagem.Caption = SolicitacaoFuncaoAutomacao.BuscaMensagem(lNSU)
            DoEvents
            Exit Do
        End If
        lbl_mensagem.Caption = SolicitacaoFuncaoAutomacao.BuscaMensagem(lNSU)
        DoEvents
        Call AguardaMS(500)
    Loop
    If SolicitacaoFuncaoAutomacao.LocalizarNSU(lNSU) Then
        lbl_mensagem.Caption = SolicitacaoFuncaoAutomacao.Mensagem
        If SolicitacaoFuncaoAutomacao.HoraAprovacao = "00:00:00" And SolicitacaoFuncaoAutomacao.HoraCancelamentoHost = "00:00:00" Then
            If SolicitacaoFuncaoAutomacao.DefineHoraCancelamentoAC(lNSU, Time) Then
                MsgBox "Tempo de aprovação/cancelamento de Serviço de Automação excedido.", vbCritical, "Tempo Excedido!"
            Else
                MsgBox "Não será possível definir cancelamento de Solicitação pela AC.", vbCritical, "Erro de Integridade!"
            End If
        Else
            If Not SolicitacaoFuncaoAutomacao.DefineHoraConfirmacaoAC(lNSU, Time) Then
                MsgBox "Não será possível definir confirmação de Solicitação pela AC.", vbCritical, "Erro de Integridade!"
            End If
        End If
    Else
        MsgBox "Não será possível localizar Solicitação de Função de Automação.", vbCritical, "Erro de Integridade!"
    End If
End Function
Private Sub AguardaProcessamentoNFCe(ByVal pNSU As Long)
    Dim xHoraInicial As Date
    Dim i As Integer
    Dim xProcessamentoConcluido As Boolean
    Dim xSegundosAAguardar As Integer

    TimerAutomacao.Enabled = False
    xProcessamentoConcluido = False

    'Define tempo em segundos pra aguardar
    xSegundosAAguardar = 120
    
    Call AguardaMS(1000)
    
    xHoraInicial = Time
    'Fica até x segundos
    Do Until DateDiff("s", xHoraInicial, Time) >= xSegundosAAguardar
        Call AguardaMS(1000)
        If MovSolicitacaoFuncaoNFe.LocalizarNSU(g_empresa, pNSU) Then
            If MovSolicitacaoFuncaoNFe.HoraAnalise_MovSolicitacaoFuncaoNFe = CDate("00:00:00") Then
                'caso a solicitação não esteja em analise ainda, verifica se petromovel está ativo
                If Not VerificarAtivarPetromovelAuto(False) Then
                    Call MovSolicitacaoFuncaoNFe.DefineHoraCancelamentoHost(MovSolicitacaoFuncaoNFe.NSU_MovSolicitacaoFuncaoNFe, Now, gVersaoSGP)
                    Exit Do
                Else
                    xHoraInicial = Time
                End If
            End If
            If MovSolicitacaoFuncaoNFe.HoraAprovacao_MovSolicitacaoFuncaoNFe <> CDate("00:00:00") Or MovSolicitacaoFuncaoNFe.HoraCancelamentoHost_MovSolicitacaoFuncaoNFe <> CDate("00:00:00") Then
                xProcessamentoConcluido = True
                Exit Do
            End If
        End If
        lbl_mensagem.Caption = SolicitacaoFuncaoAutomacao.Mensagem
        DoEvents
    Loop
    
    Call AtivaDesativaAguarde("", False)
    DoEvents
    
    'Verifica se Precisa Voltar o Abastecimento
    If g_automacao Then
        If MovSolicitacaoFuncaoNFe.CodigoRetorno_MovSolicitacaoFuncaoNFe <> 0 And MovSolicitacaoFuncaoNFe.CodigoRetorno_MovSolicitacaoFuncaoNFe <> 100 And MovSolicitacaoFuncaoNFe.CodigoRetorno_MovSolicitacaoFuncaoNFe <> 5001 Then
            LoopRetornaAbastecimentoAutomacao
        ElseIf MovSolicitacaoFuncaoNFe.HoraCancelamentoHost_MovSolicitacaoFuncaoNFe <> CDate("00:00:00") Then
            LoopRetornaAbastecimentoAutomacao
        End If
    End If
    
    If xProcessamentoConcluido = True Then
        If MovSolicitacaoFuncaoNFe.HoraAprovacao_MovSolicitacaoFuncaoNFe <> CDate("00:00:00") Then
            Dim xMensagemNFCe As String
            xMensagemNFCe = MovSolicitacaoFuncaoNFe.Mensagem_MovSolicitacaoFuncaoNFe
            MsgBox "NFCe AUTORIZADA!" & vbCrLf & "Mensagem: " & xMensagemNFCe, vbInformation, "Processamento Concluído!"
        ElseIf MovSolicitacaoFuncaoNFe.HoraCancelamentoHost_MovSolicitacaoFuncaoNFe <> CDate("00:00:00") Then
'            Dim xRsDadosParaNFCe As New adodb.Recordset
'
'            Set xRsDadosParaNFCe = ObtenhaDadosParaNFCEDocumentoEletronico(lNumeroCupom, lDataCupom)
'
'            Do Until xRsDadosParaNFCe.EOF
'                If Trim(xRsDadosParaNFCe("TipoCombustivel_MovDEItem").Value) = "" Then
'                    Dim xTipoSubEstoque As Integer
'                    xTipoSubEstoque = cboTipoSubEstoque.ItemData(cboTipoSubEstoque.ListIndex)
'
'                    If EstornaVendaProdutoNFCe(lNumeroCupom, lDataCupom, lPeriodo, xRsDadosParaNFCe("IdProduto_MovDEItem").Value, l_codigo_funcionario, xRsDadosParaNFCe("Quantidade_MovDEItem").Value, xRsDadosParaNFCe("ValorTotalLiquido_MovDEItem").Value, xTipoSubEstoque) Then
'                        If lBaixaAutomaticaNoEstoque = True Then
'                            Call AdicionaEstoque(xRsDadosParaNFCe("IdProduto_MovDEItem").Value, xRsDadosParaNFCe("Quantidade_MovDEItem").Value, xTipoSubEstoque)
'                        End If
'                    End If
'                End If
'                xRsDadosParaNFCe.MoveNext
'
'            Loop
            
            Call LoopEstornaVendaProdutoNFCeAtual
            If MovDocEletronicoCabecalho.IdClienteFornecedor > 0 Then
                Call ExcluiNotaAbastecimento
            End If
                
            MsgBox "NFCe NÃO foi Autorizada!" & vbCrLf & "Mensagem: " & MovSolicitacaoFuncaoNFe.Mensagem_MovSolicitacaoFuncaoNFe, vbCritical, "ERRO ao Processar NFCe!"
        Else
            MsgBox "Não será possível definir o processamento da NFCe.", vbCritical, "Erro de Integridade!"
        End If
    Else
        MsgBox "Tempo de solicitação de Processamento de NFCe excedido.", vbCritical, "Tempo Excedido!"
    End If
    TimerAutomacao.Enabled = True
End Sub
Private Function LoopEstornaVendaProdutoNFCeAtual() As Boolean
     LoopEstornaVendaProdutoNFCeAtual = False
     Dim xRsDadosParaNFCe As New ADODB.Recordset
    
  On Error GoTo TrataError

     Set xRsDadosParaNFCe = ObtenhaDadosParaNFCEDocumentoEletronico(lNumeroCupom, lDataCupom)
    
     Do Until xRsDadosParaNFCe.EOF
         If Trim(xRsDadosParaNFCe("TipoCombustivel_MovDEItem").Value) = "" Then
             Dim xTipoSubEstoque As Integer
             xTipoSubEstoque = xRsDadosParaNFCe("TipoSubEstoque_MovDECabecalho").Value  'cboTipoSubEstoque.ItemData(cboTipoSubEstoque.ListIndex)
    
             If EstornaVendaProdutoNFCe(lNumeroCupom, lDataCupom, lPeriodo, xRsDadosParaNFCe("IdProduto_MovDEItem").Value, l_codigo_funcionario, xRsDadosParaNFCe("Quantidade_MovDEItem").Value, xRsDadosParaNFCe("ValorTotalLiquido_MovDEItem").Value, xTipoSubEstoque) Then
                 If lBaixaAutomaticaNoEstoque = True Then
                     Call AdicionaEstoque(xRsDadosParaNFCe("IdProduto_MovDEItem").Value, xRsDadosParaNFCe("Quantidade_MovDEItem").Value, xTipoSubEstoque)
                 End If
             End If
         End If
         xRsDadosParaNFCe.MoveNext
     Loop

    LoopEstornaVendaProdutoNFCeAtual = True
    Exit Function
TrataError:
    Call CriaLogSGP("[LoopEstornaVendaProdutoNFCeAtual]", "ERRO: " & Err.Number & " - " & Err.Description, "")
End Function
Function ArquivoAutomacaoIni() As String
    lMarcaAutomacao = ReadINI("CUPOM FISCAL", "Marca da Automacao", gArquivoIni)
    If lMarcaAutomacao = "" Then
        MsgBox "Marca da automação nao informada", vbCritical, "Erro de Configuracao"
        lMarcaAutomacao = "EZTECH"
    ElseIf lMarcaAutomacao = "." Then
        ArquivoAutomacaoIni = ""
        Exit Function
    End If
    
    If lMarcaAutomacao = "COMPANY" Then
        ArquivoAutomacaoIni = "C:\Cerrado\AutoCerradoCompany\AutoCerradoCompany.INI"
    ElseIf lMarcaAutomacao = "HOROUSTECH" Then
        ArquivoAutomacaoIni = "C:\Cerrado.Net\AutoCerradoHorousTech\AutoCerradoHorousTech.INI"
    ElseIf lMarcaAutomacao = "EZTECH" Then
        ArquivoAutomacaoIni = "C:\Cerrado\AutoCerradoEZ\AutoCerradoEZ.INI"
    ElseIf lMarcaAutomacao = "IONICS" Then
        ArquivoAutomacaoIni = "C:\Cerrado\AutoCerradoIonics\AutoCerradoIonics.INI"
    End If
    If Not gArqTxt.FileExists(ArquivoAutomacaoIni) Then
        MsgBox "O Arquivo -->" & ArquivoAutomacaoIni & "<-- não foi encontrado." & vbCrLf & "O mesmo é necessário para o funcionamento completo da automação.", vbCritical, "Arquivo Inexistente!"
    End If
End Function
Private Sub AtualizaBombasAbastecimento()

On Error GoTo trata_erro
    
    For lI = 1 To lQtdBomba
        lSQL = "SELECT TOP 1 Bico, [Valor Unitario], Quantidade, ([Valor Total] - [Valor do Desconto]) AS [Valor Total], Data, Hora, [Codigo do Produto], [Tempo de Abastecimento]"
        lSQL = lSQL & "  FROM Movimento_Abastecimento"
        lSQL = lSQL & " WHERE Acerto = " & preparaBooleano(False)
        lSQL = lSQL & "   AND Bico = " & lI
        If lCaixaIndividual And g_nivel_acesso <> 1 Then
            lSQL = lSQL & " AND ( [Codigo do Funcionario] = " & l_codigo_funcionario
            lSQL = lSQL & " OR [Tempo de Abastecimento] = " & preparaTexto("11111") & " )"
        End If
        lSQL = lSQL & " ORDER BY Data DESC, Hora DESC"
        'MsgBox "AtualizaBombasAbastecimento iniciado- TESTE ALEX 1"
        Set rstAbastecimento = Conectar.RsConexao(lSQL)
        'MsgBox "AtualizaBombasAbastecimento iniciado- TESTE ALEX 2"
        If rstAbastecimento.RecordCount > 0 Then
            lAutomacaoStatusBico(lI - 1) = 6
            lbl_automacao_valor(lI - 1).Caption = Format(rstAbastecimento("Valor Total").Value, "###,##0.00")
            cmd_bico(lI - 1).Picture = LoadPicture("\VB5\SGP\ICONS\BICO_PAGAR.BMP")
            'cmd_bico(lI - 1).ToolTipText = "Bico Aguardando Pagamento."
            'É usado Produto2 para nao interferir em Produto
            If Produto2.LocalizarCodigo(rstAbastecimento("Codigo do Produto").Value) Then
                cmd_bico(lI - 1).ToolTipText = Format(rstAbastecimento("Quantidade").Value, "###,##0.00") & " Lts de " & Produto2.Nome
            End If
            lAutomacaoCodigoProduto(lI - 1) = rstAbastecimento("Codigo do Produto").Value
            lAutomacaoValorLitro(lI - 1) = rstAbastecimento("Valor Unitario").Value
            lAutomacaoLitros(lI - 1) = rstAbastecimento("Quantidade").Value
            lAutomacaoTotalAPagar(lI - 1) = rstAbastecimento("Valor Total").Value
            lAutomacaoBico(lI - 1) = rstAbastecimento("Bico").Value
            lAutomacaoData(lI - 1) = rstAbastecimento("Data").Value
            lAutomacaoHora(lI - 1) = rstAbastecimento("Hora").Value
            lAutomacaoTempoAbastecimento(lI - 1) = rstAbastecimento("Tempo de Abastecimento").Value
        Else
            lAutomacaoStatusBico(lI - 1) = 0
            lbl_automacao_valor(lI - 1).Caption = ""
            cmd_bico(lI - 1).Picture = LoadPicture("\VB5\SGP\ICONS\BICO_LIVRE.BMP")
            'cmd_bico(lI - 1).ToolTipText = "Bico Livre para Abastecimento."
        'em andamento
'            lAutomacaoStatusBico(lI - 1) = 1
'            cmd_bico(lI - 1).Picture = LoadPicture("\VB5\SGP\ICONS\BICO_ABASTECENDO.BMP")
'            cmd_bico(lI - 1).ToolTipText = "Bico com Abastecimento em Andamento."
        End If
        rstAbastecimento.Close
    Next
    Exit Sub

trata_erro:
    Call CriaLogCupom("Erro AtualizaBombasAbastecimento: Erro=" & Err.Number & " - " & Err.Description)
End Sub
Private Sub AtualizaConstantes()
    Dim xDados As String
    
    l_qtd_periodo = 1
    lQtdBomba = 8
    gQtdViasTEF = 2
    lLegislacaoPermiteIssEcf = False
    lCodigoTcsEcf = 8
    lBloqueiaEstoque = False
    lBloqueiaSubEstoque = False
    lBaixaAutomaticaNoEstoque = False
    lSerieECF = "" 'ReadINI("CUPOM FISCAL", "Serie ECF", gArquivoIni)
    
    
    
    xDados = ReadINI("CUPOM FISCAL", "Quantidade Casa Decimal", gArquivoIni)
    If Val(xDados) > 0 Then
        lEcfQtdCasasDecimais = Val(xDados)
    End If
    
'    lNumeroPDV = 1
'    'MsgBox "AtualizaConstantes iniciado- TESTE ALEX 1"
'    If ECF.LocalizarNumeroSerie(g_empresa, lSerieECF) Then
'        lNumeroPDV = ECF.Codigo
'    End If
    
    'Tasso em 23/08/2016
    lNumeroPDV = 901
    xDados = ReadINI("NFCe", "Numero do PDV", gArquivoIni)
    If Len(xDados) > 0 Then
        lNumeroPDV = Val(xDados)
    End If
    
    
    
    'MsgBox "AtualizaConstantes iniciado- TESTE ALEX 2"
    If Configuracao.LocalizarCodigo(g_empresa) Then
        gQtdViasTEF = Configuracao.QuantidadeViasTEF
        If Mid(Configuracao.OutrasConfiguracoes, 3, 1) = "S" Then
            lTEF = True
        End If
        If Mid(Configuracao.OutrasConfiguracoes, 8, 1) = "S" Then
            lLegislacaoPermiteIssEcf = True
        End If
        lCodigoTcsEcf = Mid(Configuracao.OutrasConfiguracoes, 6, 2)
        l_qtd_periodo = Configuracao.QuantidadePeriodos
        lQtdBomba = Configuracao.QuantidadeBico
        If Configuracao.ECFBaixaEstoque = True Then
            lBaixaAutomaticaNoEstoque = True
        End If
        lBloqueiaEstoque = Configuracao.BloqueiaVendaPeloEstoque
        lBloqueiaSubEstoque = Configuracao.BloqueiaVendaPeloSubEstoque
    End If
    If LiberacaoDigitacao.LocalizarCodigo(g_empresa, 2) Then
        g_cfg_data_i = LiberacaoDigitacao.DataInicial
        g_cfg_data_f = LiberacaoDigitacao.DataFinal
        g_cfg_periodo_i = LiberacaoDigitacao.PeriodoInicial
        g_cfg_periodo_f = LiberacaoDigitacao.PeriodoFinal
    End If

    lQtdMaxCombustivel = 1000
    If ConfiguracaoDiversa.LocalizarCodigo(1, "ECF: Quantidade Maxima de Combustivel") Then
        lQtdMaxCombustivel = ConfiguracaoDiversa.Valor
    End If
    lQtdMaxProduto = 100
    If ConfiguracaoDiversa.LocalizarCodigo(1, "ECF: Quantidade Maxima de Produto") Then
        lQtdMaxProduto = ConfiguracaoDiversa.Valor
    End If
    
    lBloqueiaDesconto = False
    If ConfiguracaoDiversa.LocalizarCodigo(1, "ECF: BLOQUEIA DESCONTO") Then
        lBloqueiaDesconto = ConfiguracaoDiversa.Verdadeiro
    End If
    
    lRestringeVendaCredito = 6
    If ConfiguracaoDiversa.LocalizarCodigo(1, "RESTRINGE: VENDA PELO CREDITO") Then
        lRestringeVendaCredito = ConfiguracaoDiversa.Codigo
    End If
'lRestringeVendaCredito
'lBloqueiaDesconto
    
'        ElseIf ConfiguracaoDiversa.LocalizarCodigo(1, "ECF: BLOQUEIA DESCONTO") Then
'        If txt_valor_desconto.Text > 0 Then
'            MsgBox "Empresa não configurada para desconto!", vbInformation, "Valor não aceito!"
'            txt_valor_desconto.Text = 0
'            txt_valor_desconto.SetFocus
'        Else
'            ValidaCampos2 = True
'        End If

    

    lImprimeDocumentoVinculado = True
    If ConfiguracaoDiversa.LocalizarCodigo(1, "ECF: Doc.Vinculado na Nota Abastecimento") Then
        lImprimeDocumentoVinculado = ConfiguracaoDiversa.Verdadeiro
    End If

    lQtdViasConfDiv = 1
    If ConfiguracaoDiversa.LocalizarCodigo(1, "ECF: Doc.Vinculado Qtd Vias a Imprimir") Then
        lQtdViasConfDiv = ConfiguracaoDiversa.Codigo
    End If

    lExisteMudancaHorarioVerao = False
    If MovHorarioVerao.LocalizarCodigoPendente(g_empresa, lNumeroPDV) Then
        xDados = ""
        xDados = xDados & "Bloqueio para Programação em: " & Format(MovHorarioVerao.DataParaInicioBloqueio, "dd/MM/yyyy") & " às " & Format(MovHorarioVerao.HoraParaInicioBloqueio, "HH:mm:ss")
'        If MovHorarioVerao.DataParaImpressaoReducaoZ <> CDate("00:00:00") And MovHorarioVerao.ComandoReducaoZConcluido = False Then
'            xDados = xDados & " Redução Z em: " & Format(MovHorarioVerao.DataParaImpressaoReducaoZ, "dd/MM/yyyy") & " às " & Format(MovHorarioVerao.HoraParaImpressaoReducaoZ, "HH:mm:ss")
'        Else
'            xDados = xDados & " - Redução Z Não programada"
'        End If
        xDados = xDados & " - Mudar Horário de Verão em: " & Format(MovHorarioVerao.DataParaMudancaHorario, "dd/MM/yyyy") & " às " & Format(MovHorarioVerao.HoraParaMudancaHorario, "HH:mm:ss")
        lbl_mensagem.ToolTipText = xDados
        lExisteMudancaHorarioVerao = True
    End If

    lExigeNCM = True
    If ConfiguracaoDiversa.LocalizarCodigo(1, "ECF: Exige NCM") Then
        lExigeNCM = ConfiguracaoDiversa.Verdadeiro
    End If
    
    lPOSImprimeNFCe = True
End Sub
Private Sub AtualizaRecordset(xMaxRecords As Integer)
    
    On Error GoTo FileError
    
    rst.CursorLocation = adUseClient
    rst.MaxRecords = xMaxRecords
    rst.Open lSQL, cnnSGP, adOpenForwardOnly, adLockReadOnly
    Exit Sub

FileError:
    rst.Close
    rst.CursorLocation = adUseClient
    rst.MaxRecords = xMaxRecords
    rst.Open lSQL, cnnSGP, adOpenForwardOnly, adLockReadOnly
    Exit Sub
End Sub
Private Sub AtualizaRecordset2(xMaxRecords As Integer)
    On Error GoTo FileError
    rst2.CursorLocation = adUseClient
    rst2.MaxRecords = xMaxRecords
    rst2.Open lSQL, cnnSGP, adOpenForwardOnly, adLockReadOnly
    Exit Sub
FileError:
    rst2.Close
    rst2.CursorLocation = adUseClient
    rst2.MaxRecords = xMaxRecords
    rst2.Open lSQL, cnnSGP, adOpenForwardOnly, adLockReadOnly
    Exit Sub
End Sub
Private Sub AtualizaPrecoTCS()
    Dim xArqTxt As New FileSystemObject
    Dim xArquivo As TextStream
    Dim xString As String
    Dim xSequencia As Integer
    On Error GoTo FileError
    
    
    
    Call AtivaDesativaTimer(False)
    'Set CerradoTef = Nothing
    Set CerradoTef = New CerradoComponenteTef
    
    
    xSequencia = 0
    Set xArquivo = xArqTxt.CreateTextFile("C:\TCS\TX\INTPRC.001")
    
    'Header
    xSequencia = xSequencia + 1
    xString = "H0"
    xString = xString & Format(Date, "yyyymmdd")
    xString = xString & Format(Time, "hhmmss")
    xString = xString & "000001"
    xString = xString & "INTPRC.001          "
    xString = xString & Space(102)
    xString = xString & Format(xSequencia, "000000")
    xArquivo.WriteLine (xString)
    
    
    
    'Produtos
    lSQL = "SELECT [Codigo AC], [Codigo TCS], Nome, [Preco de Venda]"
    lSQL = lSQL & "  FROM TicketCarDePara, Produto"
    lSQL = lSQL & " WHERE Produto.Codigo = TicketCarDePara.[Codigo AC]"
    lSQL = lSQL & " ORDER BY Nome"
    Set rst = Conectar.RsConexao(lSQL)
    If rst.RecordCount > 0 Then
        Do Until rst.EOF
            xSequencia = xSequencia + 1
            xString = "D1"
            xString = xString & Space(30)
            Mid(xString, 3, 5) = rst![Codigo AC]
            xString = xString & Format(rst![Codigo TCS], "00000")
            xString = xString & Format(rst![Preco de Venda] * 1000, "00000000")
            xString = xString & Space(50)
            Mid(xString, 46, 40) = rst!Nome
            xString = xString & Space(49)
            xString = xString & Format(xSequencia, "000000")
            xArquivo.WriteLine (xString)
            rst.MoveNext
        Loop
    End If
    rst.Close
    Set rst = Nothing
    
    
    'Trailer
    xSequencia = xSequencia + 1
    xString = "T0"
    xString = xString & Format(xSequencia, "000000")
    xString = xString & Space(136)
    xString = xString & Format(xSequencia, "000000")
    xArquivo.WriteLine (xString)
    xArquivo.Close
    
    If CerradoTef.SolicitacaoAlteraPrecoTCS(l_codigo_funcionario, l_nome_funcionario) Then
    Else
        MsgBox "ERRO"
    End If
    Set CerradoTef = Nothing
    Call AtivaDesativaTimer(True)
    
    Exit Sub

FileError:
    Call CriaLogCupom("Erro AtualizaPrecoTCS: Erro=" & Err.Number & " - " & Err.Description)
End Sub
Private Sub AtualizaTabelaCartaoCredito()
    Dim xDataVencimento As Date
    Dim xLocalErro As Integer
    
    Const TAMANHO_MAXIMO_NSU As Integer = 10
    
    On Error GoTo trata_erro
    
    xLocalErro = 0
    Call PreparaTipoMovimento(Produto.CodigoGrupo)
    
    xLocalErro = 1
    'lNumeroLancamentoCartao = MovCartaoCredito.ProximoRegistro(g_empresa, MovCupomFiscal.Data, CStr(MovCupomFiscal.Periodo))
    lNumeroLancamentoCartao = MovCartaoCredito.ProximoRegistro(g_empresa, MovCupomFiscal.Data)
    xLocalErro = 2
    If Not IntegracaoCaixa.LocalizarNome(g_empresa, "CARTAO " & CartaoCredito.Nome) Then
        MsgBox "Não será possível integrar com o caixa!", vbInformation, "Erro de Integridade"
    Else
        xLocalErro = 3

        If IncluiMovimentoCaixa(MovCupomFiscal.Data, MovCupomFiscal.Periodo, False, "CartaoCredito", 0, "", "") Then
            xLocalErro = 4
            'Le taxa adm do cartao
            If Not TaxaAdmCartaoCredito.LocalizarCodigo(g_empresa, CartaoCredito.Codigo) Then
                TaxaAdmCartaoCredito.TaxaCusto = CartaoCredito.TaxaCusto
                MsgBox "Taxa de Adm de Cartão de crédito não cadastrada.", vbInformation, "Erro de Integridade!"
            End If
            xLocalErro = 5
            xDataVencimento = CDate(MovCupomFiscal.Data + CartaoCredito.DiasPrazo)
            xLocalErro = 6
            MovCartaoCredito.Empresa = g_empresa
            xLocalErro = 7
            MovCartaoCredito.DataEmissao = MovCupomFiscal.Data
            xLocalErro = 8
            MovCartaoCredito.Periodo = MovCupomFiscal.Periodo
            xLocalErro = 9
            MovCartaoCredito.TipoMovimento = MovCupomFiscal.TipoMovimento
            xLocalErro = 10
            MovCartaoCredito.NumeroLancamento = lNumeroLancamentoCartao
            xLocalErro = 11
            MovCartaoCredito.CodigoCartao = lCodigoCartao
            xLocalErro = 12
            If lCartaoDataVencimento = "00:00:00" Then
                MovCartaoCredito.DataVencimento = Format(xDataVencimento, "dd/mm/yyyy")
                xLocalErro = 13
            Else
                MovCartaoCredito.DataVencimento = CDate(lCartaoDataVencimento)
                xLocalErro = 14
            End If
            MovCartaoCredito.Valor = lValorTotalUltimoCupom
            xLocalErro = 15
            MovCartaoCredito.NumeroCartao = "1"
            xLocalErro = 16
            'MovCartaoCredito.Nome = "E.C.F. " & Format(MovCupomFiscal.NumeroCupom, "###,##0")
            MovCartaoCredito.Nome = "NFCe-TEF " & Format(MovCupomFiscal.NumeroCupom, "###,##0")
            xLocalErro = 17
            MovCartaoCredito.NumeroMovimentoCaixa = MovCaixaPista.NumeroMovimento
            xLocalErro = 18
            MovCartaoCredito.TaxaAdministrativa = TaxaAdmCartaoCredito.TaxaCusto
            xLocalErro = 19
            MovCartaoCredito.NumeroIlha = lIlha
            xLocalErro = 20
            If Len(lCartaoAutorizacao) > 0 Then
                MovCartaoCredito.Autorizacao = lCartaoAutorizacao
                xLocalErro = 21
            Else
                MovCartaoCredito.Autorizacao = ""
                xLocalErro = 22
            End If
            If Val(lCartaoNSU) > 0 Then
                If Len(lCartaoNSU) >= TAMANHO_MAXIMO_NSU Then
                    MovCartaoCredito.NSU = Mid(lCartaoNSU, 6, Len(lCartaoNSU))
                    xLocalErro = 23
                Else
                    MovCartaoCredito.NSU = lCartaoNSU 'CLng(lCartaoNSU)
                    xLocalErro = 24
                End If
            Else
                MovCartaoCredito.NSU = ""
                xLocalErro = 25
            End If
            MovCartaoCredito.CodigoFuncionario = l_codigo_funcionario
            xLocalErro = 26
            
            If Not MovCartaoCredito.Incluir Then
                MsgBox "Não foi possível incluir Cartão de Crédito", vbInformation, "Erro de Integridade!"
                xLocalErro = 27
            Else
                Call CriaLogCupom("AtualizaTabelaCartaoCredito: MOVIMENTO GRAVADO OK -  MovCartaoCredito.NSU =" & MovCartaoCredito.NSU)
                xLocalErro = 28
            End If
        Else
            MsgBox "Não foi possível integrar no caixa!", vbInformation, "Erro de Integridade!"
            xLocalErro = 29
        End If
    End If
     xLocalErro = 30
    Exit Sub

trata_erro:
    Call CriaLogCupom("Erro AtualizaTabelaCartaoCredito: Erro=" & Err.Number & " - " & Err.Description & "- xLocalErro =" & xLocalErro)
    MsgBox "Erro ao tentar gravar movimento de cartão", vbCritical, "Erro de Integridade!"
End Sub
Private Function AtualizaTabelaSolicitacaoAutomacao(ByVal pTipoOperacao As String, ByVal pTexto As String) As Boolean
    Dim i As Integer
    
    On Error GoTo trata_erro
    
    AtualizaTabelaSolicitacaoAutomacao = False
    SolicitacaoFuncaoAutomacao.NSU = 1
    SolicitacaoFuncaoAutomacao.NumeroControleSolicitacao = gNumeroControleSolicitacao
    SolicitacaoFuncaoAutomacao.DataSolicitacao = Date
    SolicitacaoFuncaoAutomacao.HoraSolicitacao = Time
    SolicitacaoFuncaoAutomacao.TipoOperacao = pTipoOperacao
    SolicitacaoFuncaoAutomacao.CodigoEmpresa = g_empresa
    SolicitacaoFuncaoAutomacao.IPComputadorAC = GetIPAddress()
    SolicitacaoFuncaoAutomacao.IPInternetAC = "200??.??.??.??"
    SolicitacaoFuncaoAutomacao.SegurancaEstabelecimento = "1234"
    SolicitacaoFuncaoAutomacao.CodigoUsuario = g_usuario
    SolicitacaoFuncaoAutomacao.VersaoAC = gVersaoSGP
    SolicitacaoFuncaoAutomacao.VersaoHost = "??"
    SolicitacaoFuncaoAutomacao.Texto = pTexto
    SolicitacaoFuncaoAutomacao.HoraAnalise = CDate("00:00:00")
    SolicitacaoFuncaoAutomacao.HoraAprovacao = CDate("00:00:00")
    SolicitacaoFuncaoAutomacao.HoraCancelamentoHost = CDate("00:00:00")
    SolicitacaoFuncaoAutomacao.HoraConfirmacaoAC = CDate("00:00:00")
    SolicitacaoFuncaoAutomacao.HoraCancelamentoAC = CDate("00:00:00")
    SolicitacaoFuncaoAutomacao.Mensagem = ""
    For i = 1 To 30
        If SolicitacaoFuncaoAutomacao.Incluir Then
            AtualizaTabelaSolicitacaoAutomacao = True
            Exit For
        Else
            Call CriaLogAutomacao("AtualizaTabelaSolicitacaoAutomacao - Não foi possível gravar SolicitacaoFuncaoAutomacao - Error: " & Err.Description)
        End If
    Next
    Exit Function

trata_erro:
    Call CriaLogCupom("Erro AtualizaTabelaSolicitacaoAutomacao: Erro=" & Err.Number & " - " & Err.Description)
End Function
Private Sub AtualizaTabelaVendaProduto()
    If Not IntegracaoCaixa.LocalizarNome(g_empresa, "VENDA DE LUBRIFICANTES") Then
        MsgBox "Não será possível integrar com o caixa!", vbCritical, "Erro de Integridade!"
    Else
        If IncluiMovimentoCaixa(MovCupomFiscal.Data, MovCupomFiscal.Periodo, False, "VENDA DE LUBRIFICANTES", 0, "", "") Then
            If MovimentoLubrificante.LocalizarCodigo(g_empresa, MovCupomFiscal.Data, MovCupomFiscal.Periodo, lIlha, lTipoMovimento, MovCupomFiscal.TipoSubEstoque, MovCupomFiscal.CodigoProduto, MovCupomFiscal.operador) Then
                MovimentoLubrificante.Quantidade = MovimentoLubrificante.Quantidade + MovCupomFiscal.Quantidade
                MovimentoLubrificante.ValorTotal = MovimentoLubrificante.ValorTotal + MovCupomFiscal.ValorTotal
                If MovimentoLubrificante.Alterar(g_empresa, MovCupomFiscal.Data, MovCupomFiscal.Periodo, lIlha, lTipoMovimento, MovCupomFiscal.TipoSubEstoque, MovCupomFiscal.CodigoProduto, MovCupomFiscal.operador) Then
                Else
                    MsgBox "Não foi possível alterar o registro Venda Produto!", vbCritical, "Erro de Integridade!"
                End If
            Else
                MovimentoLubrificante.Empresa = g_empresa
                MovimentoLubrificante.Data = Format(MovCupomFiscal.Data, "dd/mm/yyyy")
                MovimentoLubrificante.Periodo = MovCupomFiscal.Periodo
                MovimentoLubrificante.NumeroIlha = lIlha
                MovimentoLubrificante.CodigoTipoSubEstoque = MovCupomFiscal.TipoSubEstoque
                MovimentoLubrificante.CodigoFuncionario = MovCupomFiscal.operador
                MovimentoLubrificante.CodigoProduto = MovCupomFiscal.CodigoProduto
                MovimentoLubrificante.Quantidade = MovCupomFiscal.Quantidade
                MovimentoLubrificante.ValorCusto = Produto.PrecoCusto
                MovimentoLubrificante.ValorVenda = MovCupomFiscal.ValorUnitario
                MovimentoLubrificante.ValorTotal = MovCupomFiscal.ValorTotal
                MovimentoLubrificante.OrdemDigitacao = 1
                MovimentoLubrificante.TipoMovimento = lTipoMovimento
                If MovimentoLubrificante.Incluir Then
                Else
                    MsgBox "Não foi possível incluir Venda de Produtos", vbCritical, "Erro de Integridade!"
                End If
            End If
        Else
            MsgBox "Não foi possível integrar no caixa!", vbCritical, "Erro de Integridade!"
        End If
    End If
End Sub
Private Function LoopIncluiMovBombaCaixa() As Boolean
    Dim i As Integer
    Dim i2 As Integer
    Dim xNome As String
    Dim xComplemento As String
    Dim xValor(0 To 5) As Currency
    Dim xNomeCombustivel(0 To 5) As String
    Dim xTipoCombustivel(0 To 5) As String
    Dim xTotal As Currency
    
    On Error GoTo trata_erro
    
    xTotal = 0
    xNome = "COMBUSTIVEIS"
    lTipoMovimento = 2
    LoopIncluiMovBombaCaixa = False
    For i = 0 To 5
        xValor(i) = 0
        xNomeCombustivel(i) = ""
        xTipoCombustivel(i) = ""
    Next
    CriaLogECF ("[LOG DE TESTE AGORA] - PASSO 03")

    For i = 1 To lQtdBomba
        If MovimentoBomba.LocalizarCodigo(g_empresa, lData, lPeriodo, i, 999) Then
            CriaLogECF ("[LOG DE TESTE AGORA] - PASSO 04 - " & i)
            CriaLogECF ("[LOG DE TESTE AGORA] - MovimentoBomba.PrecoVenda=" & MovimentoBomba.PrecoVenda & " - MovimentoBomba.QuantidadeSaida=" & MovimentoBomba.QuantidadeSaida)
            CriaLogECF ("[LOG DE TESTE AGORA] MovimentoBomba.Abertura = " & MovimentoBomba.Abertura)
            CriaLogECF ("[LOG DE TESTE AGORA] MovimentoBomba.CodigoBomba = " & MovimentoBomba.CodigoBomba)
            CriaLogECF ("[LOG DE TESTE AGORA] MovimentoBomba.Data = " & MovimentoBomba.Data)
            CriaLogECF ("[LOG DE TESTE AGORA] MovimentoBomba.Empresa = " & MovimentoBomba.Empresa)
            CriaLogECF ("[LOG DE TESTE AGORA] MovimentoBomba.Encerrante = " & MovimentoBomba.Encerrante)
            CriaLogECF ("[LOG DE TESTE AGORA] MovimentoBomba.NomeTabela = " & MovimentoBomba.NomeTabela)
            CriaLogECF ("[LOG DE TESTE AGORA] MovimentoBomba.NumeroIlha = " & MovimentoBomba.NumeroIlha)
            CriaLogECF ("[LOG DE TESTE AGORA] MovimentoBomba.NumeroTanque = " & MovimentoBomba.NumeroTanque)
            CriaLogECF ("[LOG DE TESTE AGORA] MovimentoBomba.Periodo = " & MovimentoBomba.Periodo)
            CriaLogECF ("[LOG DE TESTE AGORA] MovimentoBomba.PrecoCusto = " & MovimentoBomba.PrecoCusto)
            CriaLogECF ("[LOG DE TESTE AGORA] MovimentoBomba.PrecoVenda = " & MovimentoBomba.PrecoVenda)
            CriaLogECF ("[LOG DE TESTE AGORA] MovimentoBomba.QuantidadeSaida = " & MovimentoBomba.QuantidadeSaida)
            CriaLogECF ("[LOG DE TESTE AGORA] MovimentoBomba.SubCaixa = " & MovimentoBomba.SubCaixa)
            CriaLogECF ("[LOG DE TESTE AGORA] MovimentoBomba.TipoCombustivel = " & MovimentoBomba.TipoCombustivel)
            CriaLogECF ("[LOG DE TESTE AGORA] MovimentoBomba.TotalAcrescimo = " & MovimentoBomba.TotalAcrescimo)
            CriaLogECF ("[LOG DE TESTE AGORA] MovimentoBomba.TotalDesconto = " & MovimentoBomba.TotalDesconto)
            CriaLogECF ("[LOG DE TESTE AGORA] MovimentoBomba.TotalVendaBruto = " & MovimentoBomba.TotalVendaBruto)
            If Combustivel.LocalizarCodigo(g_empresa, MovimentoBomba.TipoCombustivel) Then
                CriaLogECF ("[LOG DE TESTE AGORA] - PASSO 05 - " & i)
                For i2 = 0 To 5
                    If xTipoCombustivel(i2) = "" Or xTipoCombustivel(i2) = MovimentoBomba.TipoCombustivel Then
                        xTipoCombustivel(i2) = MovimentoBomba.TipoCombustivel
                        xNomeCombustivel(i2) = Combustivel.Nome
                        xValor(i2) = xValor(i2) + Format(MovimentoBomba.PrecoVenda * MovimentoBomba.QuantidadeSaida, "0000000000.00")
                        
                        CriaLogECF ("[LOG DE TESTE AGORA] - PASSO 06 - xTipoCombustivel(i2)=" & xTipoCombustivel(i2) & " - xNomeCombustivel(i2)=" & xNomeCombustivel(i2) & " xValor(i2)= " & xValor(i2))

                        Exit For
                    End If
                Next
            End If
        End If
    Next
    
    
    For i = 0 To 5
        xTotal = xTotal + xValor(i)
        If xValor(i) <> 0 Then
            CriaLogECF ("[LOG DE TESTE AGORA] - PASSO 07 - " & i)
            If IntegracaoCaixa.LocalizarNome(g_empresa, "VENDA DE " & xNome) Then
                CriaLogECF ("[LOG DE TESTE AGORA] - PASSO 08 - " & i)
                
                xComplemento = Mid(xNomeCombustivel(i), 1, 27) & " Per:" & lPeriodo & " Ilha:" & lIlha & " SubCx:" & Format(999, "000")
                MovCaixaPista.Empresa = g_empresa
                MovCaixaPista.Data = lData
                MovCaixaPista.NumeroMovimento = 1
                MovCaixaPista.Valor = xValor(i)
                MovCaixaPista.CodigoHistorico = IntegracaoCaixa.HistoricoPadrao
                MovCaixaPista.DadosInterno = "BOMBA|@|999|@|"
                MovCaixaPista.CodigoLancamentoPadrao = 7
                MovCaixaPista.NumeroDocumento = ""
                MovCaixaPista.Complemento = xComplemento
                MovCaixaPista.NumeroContaDebito = IntegracaoCaixa.ContaDebito
                MovCaixaPista.NumeroContaCredito = IntegracaoCaixa.ContaCredito
                MovCaixaPista.CodigoUsuario = g_usuario
                MovCaixaPista.TipoMovimento = lTipoMovimento
                MovCaixaPista.Periodo = lPeriodo
                MovCaixaPista.NumeroIlha = lIlha
                MovCaixaPista.DataDigitacao = Format(Now, "dd/mm/yyyy")
                MovCaixaPista.HoraDigitacao = Format(Now, "HH:mm:ss")
                MovCaixaPista.DataAlteracao = "00:00:00"
                MovCaixaPista.HoraAlteracao = "00:00:00"
                If MovCaixaPista.Incluir Then
                    LoopIncluiMovBombaCaixa = True
                Else
                    LoopIncluiMovBombaCaixa = False
                End If
            Else
                CriaLogECF ("[LOG DE TESTE AGORA] - PASSO 09 - " & i & " FALHOU")
                MsgBox "Não existe a integração=" & "VENDA DE " & xNome & ".", vbCritical, "Registro Inexistente!"
            End If
        End If
    Next
    If xTotal = 0 Then
        For i = 0 To 5
            If xTipoCombustivel(i) = "G " Then
                Exit For
            End If
        Next

        CriaLogECF ("[LOG DE TESTE AGORA] - PASSO 10")
        If IntegracaoCaixa.LocalizarNome(g_empresa, "VENDA DE " & xNomeCombustivel(i)) Then
            CriaLogECF ("[LOG DE TESTE AGORA] - PASSO 10 - OK")
            CriaLogECF ("[LOG DE TESTE AGORA] - PASSO 11")
            xComplemento = Mid(xNomeCombustivel(i), 1, 27) & " Per:" & lPeriodo & " Ilha:" & lIlha & " SubCx:" & Format(999, "000")
            MovCaixaPista.Empresa = g_empresa
            MovCaixaPista.Data = lData
            MovCaixaPista.NumeroMovimento = 1
            MovCaixaPista.Valor = xValor(i)
            MovCaixaPista.CodigoHistorico = IntegracaoCaixa.HistoricoPadrao
            MovCaixaPista.DadosInterno = "BOMBA|@|999|@|"
            MovCaixaPista.CodigoLancamentoPadrao = 7
            MovCaixaPista.NumeroDocumento = ""
            MovCaixaPista.Complemento = xComplemento
            MovCaixaPista.NumeroContaDebito = IntegracaoCaixa.ContaDebito
            MovCaixaPista.NumeroContaCredito = IntegracaoCaixa.ContaCredito
            MovCaixaPista.CodigoUsuario = g_usuario
            MovCaixaPista.TipoMovimento = lTipoMovimento
            MovCaixaPista.Periodo = lPeriodo
            MovCaixaPista.NumeroIlha = lIlha
            MovCaixaPista.DataDigitacao = Format(Now, "dd/mm/yyyy")
            MovCaixaPista.HoraDigitacao = Format(Now, "HH:mm:ss")
            MovCaixaPista.DataAlteracao = "00:00:00"
            MovCaixaPista.HoraAlteracao = "00:00:00"
            If MovCaixaPista.Incluir Then
                LoopIncluiMovBombaCaixa = True
            Else
                LoopIncluiMovBombaCaixa = False
            End If
        Else
            CriaLogECF ("[LOG DE TESTE AGORA] - PASSO 10 - FALHOU - " & xNomeCombustivel(i))
            MsgBox "Não existe a integração=" & "VENDA DE " & xNome & ".", vbCritical, "Registro Inexistente!"
        End If
    End If
    Exit Function

trata_erro:
    Call CriaLogCupom("Erro LoopIncluiMovBombaCaixa: Erro=" & Err.Number & " - " & Err.Description)
End Function
'Private Sub LoopGravaCat52(ByVal pDataInicial As Date, ByVal pDataFinal As Date)
'    Dim xNumeroSerie As String
'    Dim xArqDestino As String
'    Dim xMarca As String
'    Dim xModelo As String
'    Dim xTipo As String
'    Dim xDataCat52 As Date
'    Dim i As Integer
'    Dim xExtensaoArquivo As String
'    Dim xNomeArquivo As String
'
'    'Pega Número de Série
'    xNumeroSerie = Space(21)
'    Call CriaLogCupom("Bematech_FI_NumeroSerieMFD")
'    BemaRetorno = 1 'Bematech_FI_NumeroSerieMFD(xNumeroSerie) 'NFCE - ALEX
'    xNumeroSerie = "000000000000000000001" 'NFCE - ALEX (03-08) - adicionado para corrigir erro
'    xNumeroSerie = Trim(xNumeroSerie)
'    Call CriaLogCupom("Bematech_FI_NumeroSerieMFD - BemaRetorno=" & BemaRetorno & " - NS->" & xNumeroSerie & "<-")
'    i = Len(xNumeroSerie)
'
'
'    'aqui aqui aqui erro pegar cat
'    Call CriaLogCupom("** ULTIMO CARACTER ->" & Mid(xNumeroSerie, i, 1) & "<-")
'    If Mid(xNumeroSerie, i, 1) = " " Or Asc(Mid(xNumeroSerie, i, 1)) = 32 Or Not IsNumeric(Mid(xNumeroSerie, i, 1)) Then
'        Call CriaLogCupom("** tem espaco no final do ns da ecf **")
'        Dim xteste As String
'        xteste = Mid(xNumeroSerie, 1, i - 1)
'        Call CriaLogCupom("** valor da variavel xteste->" & xteste & "<")
'        xNumeroSerie = xteste
'        Call CriaLogCupom("** valor da variavel xNumeroSerie->" & xNumeroSerie & "<")
'    Else
'        Call CriaLogCupom("** nao tem espaco no final do ns da ecf **")
'    End If
'    'Pega os últimos 5 caracteres do número de série
'    xNumeroSerie = Mid(xNumeroSerie, i - 5, 5)
'
'   'Pega Marca Modelo e Tipo do ECF
'    xMarca = Space(16)
'    xModelo = Space(21)
'    xTipo = Space(8)
'    Call CriaLogCupom("Bematech_FI_MarcaModeloTipoImpressoraMFD")
'    BemaRetorno = 1 'Bematech_FI_MarcaModeloTipoImpressoraMFD(xMarca, xModelo, xTipo) 'NFCE - ALEX
'    Call CriaLogCupom("Bematech_FI_MarcaModeloTipoImpressoraMFD - BemaRetorno=" & BemaRetorno & " - xModelo=" & xModelo)
'    xModelo = Cat52ConverteModeloBema(xModelo)
'
'    For xDataCat52 = pDataInicial To pDataFinal
'        'Verifica se existe cat52 da data específica em c:\
'        xExtensaoArquivo = Cat52ConverteDiaBema(Format(xDataCat52, "dd"))
'        xExtensaoArquivo = xExtensaoArquivo & Cat52ConverteDiaBema(Format(xDataCat52, "mm"))
'        xExtensaoArquivo = xExtensaoArquivo & Cat52ConverteDiaBema(Mid(Format(xDataCat52, "yyyy"), 3, 2))
'
'        'Faz Leitura do Cat52
'        xNomeArquivo = "C:\Cerrado.Net\Sgp\Cat52\" & xModelo & xNumeroSerie & "." & xExtensaoArquivo & ".mfd"
'        If gArqTxt.FileExists(xNomeArquivo) Then
'            Call CriaLogCupom("Arquivo Cat52 Existente:" & xNomeArquivo)
'            gStringChamada = lNumeroPDV & "|@|" & Format(xDataCat52, "dd/MM/yyyy") & "|@|" & xNomeArquivo & "|@|"
'            Call CriaLogCupom("String para grava_cat52 no banco =" & gStringChamada)
'            Call menu_personalizado.GravaSgpNetCadastroIni("grava_cat52")
'            gStringChamada = "" 'new 24/09/2015
'            Exit Sub
'        Else
'            'MsgBox "Atenção!" & vbCrLf & "Será gerado o arquivo do Cat52 da data: " & Format(xDataCat52, "dd/MM/yyyy") & "!" & vbCrLf & vbCrLf & "Este processo demora 10 minutos, favor NÃO fechar o programa!", vbInformation, "Geração do Arquivo CAT52."
'            xArqDestino = Space(512)
'            Call CriaLogCupom("Bematech_FI_GeraRegistrosCAT52MFDEx de: " & Format(xDataCat52, "dd/MM/yyyy"))
'            Call CriaLogCupom("Bematech_FI_GeraRegistrosCAT52MFDEx xNomeArquivo: " & xNomeArquivo)
'            If g_nome_empresa <> "UNIAO INFORMATICA LTDA" Then
'                Call GeraCat52(xDataCat52, xArqDestino, xNomeArquivo)
'            End If
'        End If
'
'        xDataCat52 = CDate(xDataCat52 + 1)
'    Next
'
'End Sub
Private Sub AtualTabe()
    
    On Error GoTo trata_erro
    
    lNumeroUltimoCupom = lNumeroCupom
        
    Call PreparaTipoMovimento(Produto.CodigoGrupo)
    MovCupomFiscal.Empresa = g_empresa
    MovCupomFiscal.NumeroCupom = lNumeroCupom
    MovCupomFiscal.Ordem = lOrdem
    MovCupomFiscal.Data = lData
    MovCupomFiscal.Hora = lHora
    MovCupomFiscal.DataCupom = lDataCupom
    MovCupomFiscal.Periodo = lPeriodo
    MovCupomFiscal.TipoMovimento = lTipoMovimento
    MovCupomFiscal.CodigoCliente = Val(txt_cliente.Text)
    MovCupomFiscal.CodigoConveniado = 0
    MovCupomFiscal.CodigoProduto = CLng(dtcboProduto.BoundText)
    MovCupomFiscal.ValorUnitario = fValidaValor4(txt_valor_unitario.Text)
    MovCupomFiscal.Quantidade = fValidaValor(txt_quantidade.Text)
    MovCupomFiscal.ValorTotal = fValidaValor2(txt_valor_total.Text)
    MovCupomFiscal.FormaPagamento = 0
    MovCupomFiscal.ValorRecebido = 0
    'ALEX - NFCE
'    MovCupomFiscal.NumeroCheque = ""
'    MovCupomFiscal.Telefone = ""
    MovCupomFiscal.operador = l_codigo_funcionario
    MovCupomFiscal.CupomCancelado = False
    MovCupomFiscal.ItemCancelado = False
    MovCupomFiscal.CodigoAliquota = Produto.CodigoAliquota
    MovCupomFiscal.ValorDesconto = fValidaValor2(txt_valor_desconto.Text)
    '![Valor do Acrescimo] = 0
    MovCupomFiscal.Nome = txt_nome_cliente.Text
    MovCupomFiscal.CPFCNPJ = txt_cpf.Text
    MovCupomFiscal.TipoCombustivel = Produto.TipoCombustivel
    MovCupomFiscal.CodigoECF = lNumeroPDV
    MovCupomFiscal.CodigoGrupo = Produto.CodigoGrupo
    MovCupomFiscal.TipoSubEstoque = cboTipoSubEstoque.ItemData(cboTipoSubEstoque.ListIndex)
    MovCupomFiscal.ValorDescontoEmbutido = 0
'ALEX - NFCE
'    If Val(txt_numero_nota_abastecimento.Text) > 0 Then
'        MovCupomFiscal.NumeroCheque = CLng(txt_numero_nota_abastecimento.Text)
'    Else
'        MovCupomFiscal.NumeroCheque = 0
'    End If
    Exit Sub

trata_erro:
    Call CriaLogCupom("Erro AtualTabe: Erro=" & Err.Number & " - " & Err.Description)
End Sub
Private Sub AtualTela()

    On Error GoTo trata_erro
    
    Dim i As Integer
    
    lNumeroUltimoCupom = MovCupomFiscal.NumeroCupom
    lNumeroCupom = MovCupomFiscal.NumeroCupom
    lData = MovCupomFiscal.Data
    lOrdem = MovCupomFiscal.Ordem
    lNumeroCupom = MovCupomFiscal.NumeroCupom
    lOrdem = MovCupomFiscal.Ordem
    lHora = MovCupomFiscal.Hora
    lPeriodo = MovCupomFiscal.Periodo
    cboTipoSubEstoque.ListIndex = -1
    For i = 0 To cboTipoSubEstoque.ListCount - 1
        If cboTipoSubEstoque.ItemData(i) = MovCupomFiscal.TipoSubEstoque Then
            cboTipoSubEstoque.ListIndex = i
            Exit For
        End If
    Next
    txt_cliente.Text = MovCupomFiscal.CodigoCliente
    dtcboCliente.BoundText = MovCupomFiscal.CodigoCliente
    txt_produto.Text = MovCupomFiscal.CodigoProduto
    If Produto.LocalizarCodigo(MovCupomFiscal.CodigoProduto) Then
        dtcboProduto.BoundText = MovCupomFiscal.CodigoProduto
        If Not Aliquota.LocalizarCodigo(lSerieECF, Produto.CodigoAliquota) Then
            MsgBox "Aliquota não cadastrada!", vbInformation, "Erro de Integridade!"
        End If
    Else
        dtcboProduto.BoundText = ""
    End If
    txt_valor_unitario.Text = Format(MovCupomFiscal.ValorUnitario, "###,##0.0000")
    txt_quantidade.Text = Format(MovCupomFiscal.Quantidade, "###,##0.000")
    txt_valor_total.Text = Format(MovCupomFiscal.ValorTotal, "###,##0.00")
    VerificaLiberacaoDigitacao
    Exit Sub

trata_erro:
    Call CriaLogCupom("Erro AdicionaEstoque: Erro=" & Err.Number & " - " & Err.Description)
End Sub
Private Sub BuscaPeriodo()
    If LiberacaoDigitacao.LocalizarCodigo(g_empresa, 2) Then
        g_cfg_periodo_i = LiberacaoDigitacao.PeriodoInicial
        g_cfg_periodo_f = LiberacaoDigitacao.PeriodoFinal
        g_cfg_data_i = LiberacaoDigitacao.DataInicial
        g_cfg_data_f = LiberacaoDigitacao.DataFinal
    End If
    If PeriodoTrocaOleo.LocalizarCodigo(g_empresa, Val(txt_funcionario_ponto.Text)) Then
        g_cfg_periodo_i = PeriodoTrocaOleo.Periodo
        g_cfg_periodo_f = PeriodoTrocaOleo.Periodo
        lTipoMovimento = 3
        cboTipoSubEstoque.ListIndex = lTipoMovimento - 2
    End If
    lPeriodo = g_cfg_periodo_i
End Sub
Function BuscaRegistro(x_numero_cupom As Long, x_data As Date, x_ordem As Integer) As Boolean
    BuscaRegistro = False
    If MovCupomFiscal.LocalizarCodigo(g_empresa, lNumeroPDV, x_numero_cupom, x_data, x_ordem) Then
        BuscaRegistro = True
    End If
End Function
Private Sub LoopRetornaAbastecimentoAutomacao()

    On Error GoTo trata_erro
    
    lSQL = "SELECT Data, Hora, Bico"
    lSQL = lSQL & "  FROM Movimento_Abastecimento"
    lSQL = lSQL & " WHERE Acerto = " & preparaBooleano(True)
    lSQL = lSQL & "   AND [Numero do Cupom] = " & lNumeroCupom
    lSQL = lSQL & " ORDER BY Data ASC, Hora ASC, Bico ASC"
    Set rstAbastecimento = Conectar.RsConexao(lSQL)
    If rstAbastecimento.RecordCount > 0 Then
        Do Until rstAbastecimento.EOF
            If MovimentoAbastecimento.LocalizarCodigo(g_empresa, rstAbastecimento("Data").Value, rstAbastecimento("Hora").Value, rstAbastecimento("Bico").Value) Then
                MovimentoAbastecimento.Acerto = False
                MovimentoAbastecimento.NumeroCupom = 0
                MovimentoAbastecimento.CodigoECF = 0
                MovimentoAbastecimento.DocumentoGerado = ""
                If Not MovimentoAbastecimento.Alterar(g_empresa, rstAbastecimento("Data").Value, rstAbastecimento("Hora").Value, rstAbastecimento("Bico").Value) Then
                    MsgBox "Não foi possível alterar o abastecimento!", vbInformation, "Erro de Integridade!"
                End If
            Else
                MsgBox "Não foi possível localizar abastecimento!", vbInformation, "Erro de Integridade!"
            End If
            rstAbastecimento.MoveNext
        Loop
    End If
    rstAbastecimento.Close
    Exit Sub

trata_erro:
    Call CriaLogAutomacao("Erro LoopRetornaAbastecimentoAutomacao: Erro=" & Err.Number & " - " & Err.Description)
End Sub
Private Sub AutomacaoAlteraTabeAbastecimento(ByVal pIndiceBico As Integer, ByVal pNumeroCupom As Long)

    On Error GoTo trata_erro
    
    Call CriaLogAutomacao("Movimento_Nfce_Auto.[AutomacaoAlteraTabeAbastecimento] Rotina Iniciada - pIndiceBico=" & pIndiceBico & " - pNumeroCupom=" & pNumeroCupom)

    If MovimentoAbastecimento.LocalizarCodigo(g_empresa, lAutomacaoData(pIndiceBico), lAutomacaoHora(pIndiceBico), lAutomacaoBico(pIndiceBico)) Then
    
        Call CriaLogAutomacao("Movimento_Nfce_Auto.[AutomacaoAlteraTabeAbastecimento] ABASTECIMENTO LOCALIZADO - lAutomacaoData(pIndiceBico)=" & lAutomacaoData(pIndiceBico) & " - lAutomacaoHora(pIndiceBico)=" & lAutomacaoHora(pIndiceBico) & " - lAutomacaoBico(pIndiceBico)=" & lAutomacaoBico(pIndiceBico))
        
        MovimentoAbastecimento.Acerto = True
        MovimentoAbastecimento.NumeroCupom = pNumeroCupom
        MovimentoAbastecimento.CodigoECF = lNumeroPDV
        MovimentoAbastecimento.DocumentoGerado = "NFCe"
        'CF   - Cupom Fiscal
        'NT   - Nota Abastecimento
        'CP   - Cupom Complementar
        'AF   - Afericao
        'CHVIS- Cheque A Vista
        'CHPRE- Cheque Pre-Datado
        'CRT  - Cartao de Credito
        'DIN  - Dinheiro
        'DESPC- Despesa de Caixa
        'VALEF- Vale de Funcionario
        'VLABR- Vale Abastecimento Recebido
        'CRAR - Credito Antecipado Recebido
        'NFCe - Nota Fiscal do Consumidor Eletronica
        If Not MovimentoAbastecimento.Alterar(g_empresa, lAutomacaoData(pIndiceBico), lAutomacaoHora(pIndiceBico), lAutomacaoBico(pIndiceBico)) Then
            Call CriaLogAutomacao("Movimento_Nfce_Auto.[AutomacaoAlteraTabeAbastecimento] NÃO FOI POSSÍVEL BAIXAR O ABASTECIMENTO - lAutomacaoData(pIndiceBico)=" & lAutomacaoData(pIndiceBico) & " - lAutomacaoHora(pIndiceBico)=" & lAutomacaoHora(pIndiceBico) & " - lAutomacaoBico(pIndiceBico)=" & lAutomacaoBico(pIndiceBico))
            MsgBox "Não foi possível alterar o abastecimento!", vbInformation, "Erro de Integridade!"
        Else
            Call CriaLogAutomacao("Movimento_Nfce_Auto.[AutomacaoAlteraTabeAbastecimento] Abastecimento Baixado com Sucesso - lAutomacaoData(pIndiceBico)=" & lAutomacaoData(pIndiceBico) & " - lAutomacaoHora(pIndiceBico)=" & lAutomacaoHora(pIndiceBico) & " - lAutomacaoBico(pIndiceBico)=" & lAutomacaoBico(pIndiceBico))
        End If
    Else
        Call CriaLogAutomacao("Movimento_Nfce_Auto.[AutomacaoAlteraTabeAbastecimento] ABASTECIMENTO NÃO LOCALIZADO - lAutomacaoData(pIndiceBico)=" & lAutomacaoData(pIndiceBico) & " - lAutomacaoHora(pIndiceBico)=" & lAutomacaoHora(pIndiceBico) & " - lAutomacaoBico(pIndiceBico)=" & lAutomacaoBico(pIndiceBico))
        MsgBox "Não foi possível localizar abastecimento!", vbInformation, "Erro de Integridade!"
    End If
    Exit Sub

trata_erro:
    Call CriaLogAutomacao("Erro AutomacaoAlteraTabeAbastecimento: Erro=" & Err.Number & " - " & Err.Description)
End Sub
Private Sub AutomacaoAlteraArredondamento(ByVal pData As Date, ByVal pHora As Date, ByVal pBico As Integer, ByVal pIndex As Integer)
    Dim xDiferenca As Currency
    Dim xValorTotal As Currency

    On Error GoTo trata_erro
    
    If MovimentoAbastecimento.LocalizarCodigo(g_empresa, pData, pHora, pBico) Then
        xValorTotal = Round(MovimentoAbastecimento.Quantidade * MovimentoAbastecimento.ValorUnitario, 2)
        xDiferenca = MovimentoAbastecimento.ValorTotal - xValorTotal
        If xDiferenca <> 0 Then
            MovimentoAbastecimento.Quantidade = Round(MovimentoAbastecimento.ValorTotal / MovimentoAbastecimento.ValorUnitario, 3)
            If pIndex <> -1 Then
                lAutomacaoLitros(pIndex) = MovimentoAbastecimento.Quantidade
            End If
            If Not MovimentoAbastecimento.Alterar(g_empresa, pData, pHora, pBico) Then
                MsgBox "Erro em: AutomacaoAlteraArredondamento" & vbCrLf & "Não foi possível alterar o abastecimento!", vbInformation, "Erro de Integridade!"
            End If
        End If
    Else
        MsgBox "Erro em: AutomacaoAlteraArredondamento" & vbCrLf & "Não foi possível localizar abastecimento!", vbInformation, "Erro de Integridade!"
    End If
    Exit Sub

trata_erro:
    Call CriaLogAutomacao("Erro AutomacaoAlteraArredondamento: Erro=" & Err.Number & " - " & Err.Description)
End Sub
Private Sub AutomacaoInicio()
    For lI = 0 To (lQtdBomba - 1)
        lAutomacaoStatusBico(lI) = 0
        lbl_automacao_valor(lI).Caption = ""
        lAutomacaoCodigoProduto(lI) = 0
        lAutomacaoBico(lI) = 0
        lAutomacaoData(lI) = 0
        lAutomacaoHora(lI) = 0
        lAutomacaoTempoAbastecimento(lI) = ""
        lAutomacaoValorLitro(lI) = 0
        lAutomacaoLitros(lI) = 0
        lAutomacaoTotalAPagar(lI) = 0
    Next
    lAutomacaoFlag = 0
    TimerAutomacao.Interval = 1000
    TimerAutomacao.Enabled = True
End Sub
Private Sub AutomacaoMostraBicos()
    Dim i As Integer
    Dim xPassoDebugarErro As String
    
    On Error GoTo FileError
    
    xPassoDebugarErro = "1"
    For i = 0 To (lQtdBomba - 1)
        If lAutomacaoStatusBico(i) = 0 Then
            xPassoDebugarErro = "2"
            cmd_bico(i).Picture = LoadPicture("\VB5\SGP\ICONS\BICO_LIVRE.BMP")
            cmd_bico(i).ToolTipText = "Bico Livre para Abastecimento."
        ElseIf lAutomacaoStatusBico(i) = 1 Then
            xPassoDebugarErro = "3"
            cmd_bico(i).Picture = LoadPicture("\VB5\SGP\ICONS\BICO_ABASTECENDO.BMP")
            cmd_bico(i).ToolTipText = "Bico com Abastecimento em Andamento."
        ElseIf lAutomacaoStatusBico(i) = 6 Then
            xPassoDebugarErro = "4"
            cmd_bico(i).Picture = LoadPicture("\VB5\SGP\ICONS\BICO_PAGAR.BMP")
            cmd_bico(i).ToolTipText = "Bico Aguardando Pagamento."
        End If
        If EncerranteAtual.LocalizarCodigo(g_empresa, i + 1) Then
            If EncerranteAtual.Situacao = "INVALIDA" Then
                cmd_bico(i).Picture = LoadPicture("\VB5\SGP\ICONS\BICO_ERRO.BMP")
                cmd_bico(i).ToolTipText = "Bico inválido."
            ElseIf EncerranteAtual.Situacao = "NAO RESPON" Then
                cmd_bico(i).Picture = LoadPicture("\VB5\SGP\ICONS\BICO_ERRO.BMP")
                cmd_bico(i).ToolTipText = "Bico NÃO respondendo."
            ElseIf EncerranteAtual.Situacao = "AUTORIZADA" And lAutomacaoStatusBico(i) <> 6 Then
                cmd_bico(i).Picture = LoadPicture("\VB5\SGP\ICONS\BICO_LIVRE.BMP")
                cmd_bico(i).ToolTipText = "Bico autorizado para Abastecimento (Livre)."
            ElseIf EncerranteAtual.Situacao = "INICIADA" Then
                cmd_bico(i).Picture = LoadPicture("\VB5\SGP\ICONS\BICO_ABASTECENDO.BMP")
                cmd_bico(i).ToolTipText = "Bico iniciando o Abastecimento."
            ElseIf EncerranteAtual.Situacao = "ABASTECEND" Then
                cmd_bico(i).Picture = LoadPicture("\VB5\SGP\ICONS\BICO_ABASTECENDO2.BMP")
                cmd_bico(i).ToolTipText = "Bico com Abastecimento em Andamento."
                lbl_automacao_valor(i).Caption = Format(EncerranteAtual.Litragem, "###,##0.00")
            ElseIf EncerranteAtual.Situacao = "CONCLUIDA" Or EncerranteAtual.Situacao = "CONCLUIDO" Then
                cmd_bico(i).Picture = LoadPicture("\VB5\SGP\ICONS\BICO_PAGAR.BMP")
            ElseIf EncerranteAtual.Situacao = "FINALIZADO" Then
                cmd_bico(i).Picture = LoadPicture("\VB5\SGP\ICONS\BICO_PAGAR.BMP")
            End If
        End If
    Next
    Exit Sub

FileError:
    Call CriaLogAutomacao(Time & " ERRO AutomacaoMostraBicos: i=" & i & " - " & Error & " - xPassoDebugarErro:" & xPassoDebugarErro & " - Err: " & Err)
    Exit Sub
End Sub
Private Sub BaixaAbastecimentoAcertado()
    On Error GoTo trata_erro
    Dim xGravaBaixa As Boolean
    
    lI = 0
    lSQL = "SELECT TOP 10000 Data, Hora, Bico"
    lSQL = lSQL & "  FROM Movimento_Abastecimento"
    lSQL = lSQL & " WHERE Acerto = " & preparaBooleano(True)
    'lSQL = lSQL & "   AND Data = '26/05/2015'"
    lSQL = lSQL & "   AND Data < " & preparaData(Date)
    lSQL = lSQL & " ORDER BY Data ASC, Hora ASC"
    
    
'    lSQL = "SELECT TOP 1 Data, Hora, Bico"
'    lSQL = lSQL & "  FROM Movimento_Abastecimento"
'    lSQL = lSQL & " WHERE Acerto = " & preparaBooleano(True)
'    lSQL = lSQL & "   AND Data = '16/07/2015'"
'    'lSQL = lSQL & "   AND Data < " & preparaData(Date)
'    lSQL = lSQL & " ORDER BY Data ASC, Hora ASC"
    
    
    
    Set rstAbastecimento = Conectar.RsConexao(lSQL)
    If rstAbastecimento.RecordCount > 0 Then
        Do Until rstAbastecimento.EOF
            lI = lI + 1
            If Not IsNull(rstAbastecimento("Hora").Value) Then
                If MovimentoAbastecimento.LocalizarCodigo(g_empresa, rstAbastecimento("Data").Value, rstAbastecimento("Hora").Value, rstAbastecimento("Bico").Value) Then
                    xGravaBaixa = True
                    Call CriaLogCupom("BaixaAbastecimentoAcertado: Lendo registro na Baixa. Data=" & rstAbastecimento("Data").Value & " - Hora=" & rstAbastecimento("Hora").Value & " - Bico=" & rstAbastecimento("Bico").Value)
                    If BaixaAbastecimento.LocalizarCodigo(g_empresa, rstAbastecimento("Bico").Value, rstAbastecimento("Data").Value, rstAbastecimento("Hora").Value) Then
                        Call CriaLogCupom("BaixaAbastecimentoAcertado: Registro encontrado na baixa Data=" & rstAbastecimento("Data").Value & " - Hora=" & rstAbastecimento("Hora").Value & " - Bico=" & rstAbastecimento("Bico").Value)
                        Call CriaLogCupom("BaixaAbastecimentoAcertado: Registro encontrado na baixa BaixaAbastecimento.Quantidade=" & BaixaAbastecimento.Quantidade & " - MovimentoAbastecimento.Quantidade=" & MovimentoAbastecimento.Quantidade)
                        If BaixaAbastecimento.Quantidade = MovimentoAbastecimento.Quantidade Then
                            xGravaBaixa = False
                            If Not MovimentoAbastecimento.Excluir(g_empresa, BaixaAbastecimento.Data, BaixaAbastecimento.Hora, BaixaAbastecimento.Bico) Then
                                Call CriaLogCupom("Erro BaixaAbastecimentoAcertado: Baixa Existente. Erro ao excluir abastecimento. Data=" & BaixaAbastecimento.Data & " - Hora=" & BaixaAbastecimento.Hora & " - Bico=" & BaixaAbastecimento.Bico)
                            End If
                        Else
                            Call CriaLogCupom("Erro BaixaAbastecimentoAcertado: Baixa Existente e quantidade diferente. BaixaAbastecimento.Quantidade=" & BaixaAbastecimento.Quantidade & " - MovimentoAbastecimento.Quantidade=" & MovimentoAbastecimento.Quantidade & " - Hora=" & BaixaAbastecimento.Hora & " - Bico=" & BaixaAbastecimento.Bico)
                        End If
                    End If
                    If xGravaBaixa = True Then
                        BaixaAbastecimento.Empresa = MovimentoAbastecimento.Empresa
                        BaixaAbastecimento.Bico = MovimentoAbastecimento.Bico
                        BaixaAbastecimento.Data = MovimentoAbastecimento.Data
                        BaixaAbastecimento.Hora = MovimentoAbastecimento.Hora
                        BaixaAbastecimento.TempoAbastecimento = MovimentoAbastecimento.TempoAbastecimento
                        BaixaAbastecimento.CodigoProduto = MovimentoAbastecimento.CodigoProduto
                        BaixaAbastecimento.ValorUnitario = MovimentoAbastecimento.ValorUnitario
                        BaixaAbastecimento.Quantidade = MovimentoAbastecimento.Quantidade
                        BaixaAbastecimento.ValorTotal = MovimentoAbastecimento.ValorTotal
                        BaixaAbastecimento.NumeroAbastecimento = MovimentoAbastecimento.NumeroAbastecimento
                        BaixaAbastecimento.Encerrante = MovimentoAbastecimento.Encerrante
                        If Len(BaixaAbastecimento.Encerrante) > 10 Then
                            BaixaAbastecimento.Encerrante = Mid(BaixaAbastecimento.Encerrante, 1, 10)
                        End If
                        BaixaAbastecimento.StringAutomacao = MovimentoAbastecimento.StringAutomacao
                        BaixaAbastecimento.CodigoECF = MovimentoAbastecimento.CodigoECF
                        BaixaAbastecimento.NumeroCupom = MovimentoAbastecimento.NumeroCupom
                        BaixaAbastecimento.DataBaixa = Date
                        BaixaAbastecimento.HoraBaixa = Time
                        BaixaAbastecimento.DocumentoGerado = MovimentoAbastecimento.DocumentoGerado
                        BaixaAbastecimento.ComplementoDocumentoGerado = MovimentoAbastecimento.ComplementoDocumentoGerado
                        BaixaAbastecimento.ValorDesconto = MovimentoAbastecimento.ValorDesconto
                        BaixaAbastecimento.Acerto = MovimentoAbastecimento.Acerto
                        BaixaAbastecimento.EncerranteInicial = MovimentoAbastecimento.EncerranteInicial
                        BaixaAbastecimento.TipoCombustivel = MovimentoAbastecimento.TipoCombustivel
                        BaixaAbastecimento.CodigoFuncionario = MovimentoAbastecimento.CodigoFuncionario
                        BaixaAbastecimento.Periodo = MovimentoAbastecimento.Periodo
                        BaixaAbastecimento.ValorDescontoNivelPreco = MovimentoAbastecimento.ValorDescontoNivelPreco
                        BaixaAbastecimento.ValorAcrescimoNivelPreco = MovimentoAbastecimento.ValorAcrescimoNivelPreco
                        
                        If BaixaAbastecimento.Incluir Then
                            If Not MovimentoAbastecimento.Excluir(g_empresa, BaixaAbastecimento.Data, BaixaAbastecimento.Hora, BaixaAbastecimento.Bico) Then
                                Call CriaLogCupom("Erro BaixaAbastecimentoAcertado : Erro ao excluir abastecimento. Data=" & BaixaAbastecimento.Data & " - Hora=" & BaixaAbastecimento.Hora & " - Bico=" & BaixaAbastecimento.Bico)
                            End If
                        Else
                            Call CriaLogCupom("Erro BaixaAbastecimentoAcertado: Erro ao incluir baixa de abastecimento. Data=" & BaixaAbastecimento.Data & " - Hora=" & BaixaAbastecimento.Hora & " - Bico=" & BaixaAbastecimento.Bico)
                        End If
                    End If
                Else
                    Call CriaLogCupom("Erro BaixaAbastecimentoAcertado: abastecimento não encontrado. Data=" & rstAbastecimento("Data").Value & " - Hora=" & rstAbastecimento("Hora").Value & " - Bico=" & rstAbastecimento("Bico").Value)
                End If
            End If
            rstAbastecimento.MoveNext
        Loop
    End If
    rstAbastecimento.Close
    Exit Sub

trata_erro:
    Call CriaLogCupom("Erro BaixaAbastecimentoAcertado: Erro=" & Err.Number & " - " & Err.Description)
End Sub
Function BuscaDados() As Boolean
    'MsgBox "BuscaDados iniciado- TESTE ALEX"
    BuscaDados = False
    If MovCupomFiscal.LocalizarUltimo(g_empresa, lNumeroPDV) Then
        BuscaDados = True
        Call MontaCupomVideo(MovCupomFiscal.NumeroCupom, MovCupomFiscal.Data)
    Else
        LimpaTela
    End If
    'MsgBox "BuscaDados finalizado- TESTE ALEX"
End Function
Function ComunicaAutomacaoCerradoArq(ByVal pComando As String, ByVal pParametro As String) As Boolean
    Dim xArquivoTmp As String
    Dim xArquivoPedido As String
    Dim xArquivoResp As String
    Dim xComputadorAutomacao As String
    Dim xPastaAutomacao As String
    Dim xHoraInicial As Date
    Dim xComando As String
    Dim xRetorno As String
    Dim xParametro As String
    Dim xTempo As Integer
'    Dim xNumeroEmailInicial As String

    On Error GoTo FileError

    ComunicaAutomacaoCerradoArq = False
    
'    'Grava tipo de email
'    xNumeroEmailInicial = gNumeroEmailInicial
'    Call WriteINI("EMAIL", "Numero do Email", xNumeroEmailInicial, lNomeArquivoAutomacaoIni)
'    If gNumeroEmailInicial > 0 Then
'        Call WriteINI("EMAIL", "Concluido", "NAO", lNomeArquivoAutomacaoIni)
'    Else
'        Call WriteINI("EMAIL", "Concluido", "SIM", lNomeArquivoAutomacaoIni)
'    End If
    
    'Pega o NOME do computador que tem Ligado Fisicamente o Equipamento de Automação
    xComputadorAutomacao = ReadINI("LOCALIZACAO", "Computador com automacao", lNomeArquivoAutomacaoIni)
    xPastaAutomacao = ReadINI("LOCALIZACAO", "Pasta da automacao", lNomeArquivoAutomacaoIni)
    If xPastaAutomacao = "" Then
        xPastaAutomacao = "Automacao"
    End If
    If xComputadorAutomacao = "" Then
        xComputadorAutomacao = GetIPHostName()
    End If
    
    'Monta nome do Computador + Diretório + Arquivo do Pedido de comunicação
    'Ex: \\Servidor\Automacao\Pedido_ddmmyyyy_HHmmss.TMP
    xArquivoTmp = "\\" & xComputadorAutomacao & "\" & xPastaAutomacao & "\Pedido_" & Format(Date, "ddmmyyyy") & "_" & Format(Time, "HHmmss") & ".TMP"
    xArquivoPedido = Mid(xArquivoTmp, 1, Len(xArquivoTmp) - 3) & "AUT"

    'Cria o arquivo .TMP de comunicação
    Set gArquivoTXT = gArqTxt.CreateTextFile(xArquivoTmp)
    gArquivoTXT.WriteLine ("[PEDIDO AUTOMACAO]")
    gArquivoTXT.WriteLine ("Comando=" & pComando)
    gArquivoTXT.WriteLine ("Origem=" & GetIPHostName())
    gArquivoTXT.WriteLine ("Parametro=" & pParametro)
    gArquivoTXT.Close

    'Renomeia arquivo .TMP para .AUT
    If gArqTxt.FileExists(xArquivoTmp) Then
        gArqTxt.MoveFile (xArquivoTmp), (xArquivoPedido)
    End If
    
    'Monta nome do Arquivo de Retorno
    xArquivoResp = "C:\" & xPastaAutomacao & "\Retorno_" & Mid(xArquivoPedido, Len(xArquivoPedido) - 19 + 1, 19)

    xTempo = 7
    If pComando = "AUTOMACAO LEITURA ENCERRANTE" Then
        xTempo = lQtdBomba * 10
    End If
    If pComando = "AUTOMACAO INCLUI CARTAO" Then
        xTempo = 45
    End If
'    If xTempo > 80 Then
'        xTempo = 30
'    End If
    'Aguarda até 7 Segundos para o retorno
    xHoraInicial = Time
    Do Until DateDiff("s", xHoraInicial, Time) >= xTempo
        If gArqTxt.FileExists(xArquivoResp) Then
            Exit Do
        End If
        DoEvents
    Loop
    
    'Verifica se o Retorno existe
    If gArqTxt.FileExists(xArquivoResp) Then
        'Existindo lê o retorno
        xComando = ReadINI("RETORNO AUTOMACAO", "Comando", xArquivoResp)
        xRetorno = ReadINI("RETORNO AUTOMACAO", "Retorno", xArquivoResp)
        xParametro = ReadINI("RETORNO AUTOMACAO", "Parametro", xArquivoResp)
        'Deleta o arquivo de retorno retorno
        gArqTxt.DeleteFile (xArquivoResp)
        If xRetorno = "OK" Then
            If pComando = "AUTOMACAO INCLUI CARTAO" Then
                CartaoAbastecimento.PosicaoRegistro = fRetiraString(xParametro, 1)
                If Len(CartaoAbastecimento.PosicaoRegistro) = 6 Then
                    If CartaoAbastecimento.Alterar(g_empresa, CartaoAbastecimento.NumeroCartao) Then
                        ComunicaAutomacaoCerradoArq = True
                        MsgBox "Cartão Incluído no IdentFID!", vbInformation + vbOKOnly, "Cartão Incluído!"
                    Else
                        ComunicaAutomacaoCerradoArq = False
                        MsgBox "Erro ao alterar o Registro do Cartão de Abastecimento!", vbOKOnly + vbCritical, "Cartão Não Alterado!"
                    End If
                Else
                    MsgBox "Erro ao criticar retorno da Inclusão de Cartão de Abastecimento!" & vbCrLf & "Posicao do Registro=" & CartaoAbastecimento.PosicaoRegistro, vbCritical, "Cartão Não Incluído!"
                    ComunicaAutomacaoCerradoArq = False
                End If
            Else
                ComunicaAutomacaoCerradoArq = True
            End If
        End If
    Else
        'MsgBox "arquivo nao encontrado=" & xArquivoResp
        'Deleta o Arquivo de pedido
        'Pois fica sub-entendido que o mesmo ainda existe
        gArqTxt.DeleteFile (xArquivoPedido)
    End If
    
    Exit Function

FileError:
    Call CriaLogCupom("Erro ComunicaAutomacaoCerradoArq: Erro=" & Err.Number & " - " & Err.Description)
    MsgBox "Erro ao tentar comunicação com o programa AutoCerrado!", vbCritical, "Erro na Automação de Bomba!"
    Exit Function
End Function
Function ComunicaAutomacaoCerradoBD(ByVal pComando As String, ByVal pParametro As String) As Boolean
    'Dim xNumeroEmailInicial As String
    Dim xTextoComando As String
    Dim xTempo As Integer
    Dim xHoraInicial As Date
'    Dim xArquivoTmp As String
'    Dim xArquivoPedido As String
'    Dim xArquivoResp As String
'    Dim xComputadorAutomacao As String
'    Dim xPastaAutomacao As String
'    Dim xComando As String
'    Dim xRetorno As String
'    Dim xParametro As String

    On Error GoTo FileError

    ComunicaAutomacaoCerradoBD = False
    
'    'Grava tipo de email
'    xNumeroEmailInicial = gNumeroEmailInicial
'    Call WriteINI("EMAIL", "Numero do Email", xNumeroEmailInicial, lNomeArquivoAutomacaoIni)
'    If gNumeroEmailInicial > 0 Then
'        Call WriteINI("EMAIL", "Concluido", "NAO", lNomeArquivoAutomacaoIni)
'    Else
'        Call WriteINI("EMAIL", "Concluido", "SIM", lNomeArquivoAutomacaoIni)
'    End If
    
'    'Pega o NOME do computador que tem Ligado Fisicamente o Equipamento de Automação
'    xComputadorAutomacao = ReadINI("LOCALIZACAO", "Computador com automacao", lNomeArquivoAutomacaoIni)
'    xPastaAutomacao = ReadINI("LOCALIZACAO", "Pasta da automacao", lNomeArquivoAutomacaoIni)
'    If xPastaAutomacao = "" Then
'        xPastaAutomacao = "Automacao"
'    End If
'    If xComputadorAutomacao = "" Then
'        xComputadorAutomacao = GetIPHostName()
'    End If
    
'    'Monta nome do Computador + Diretório + Arquivo do Pedido de comunicação
'    'Ex: \\Servidor\Automacao\Pedido_ddmmyyyy_HHmmss.TMP
'    xArquivoTmp = "\\" & xComputadorAutomacao & "\" & xPastaAutomacao & "\Pedido_" & Format(Date, "ddmmyyyy") & "_" & Format(Time, "HHmmss") & ".TMP"
'    xArquivoPedido = Mid(xArquivoTmp, 1, Len(xArquivoTmp) - 3) & "AUT"

'    'Cria o arquivo .TMP de comunicação
'    Set gArquivoTXT = gArqTxt.CreateTextFile(xArquivoTmp)
'    gArquivoTXT.WriteLine ("[PEDIDO AUTOMACAO]")
'    gArquivoTXT.WriteLine ("Comando=" & pComando)
'    gArquivoTXT.WriteLine ("Origem=" & GetIPHostName())
'    gArquivoTXT.WriteLine ("Parametro=" & pParametro)
'    gArquivoTXT.Close
    
    

'    'Renomeia arquivo .TMP para .AUT
'    If gArqTxt.FileExists(xArquivoTmp) Then
'        gArqTxt.MoveFile (xArquivoTmp), (xArquivoPedido)
'    End If
'
'    'Monta nome do Arquivo de Retorno
'    xArquivoResp = "C:\" & xPastaAutomacao & "\Retorno_" & Mid(xArquivoPedido, Len(xArquivoPedido) - 19 + 1, 19)

    xTempo = 7
    If pComando = "AUTOMACAO LEITURA ENCERRANTE" Then
        xTempo = lQtdBomba * 10
    End If
    If pComando = "INCLUI CARTAO RFID" Then
        xTempo = 45
    ElseIf pComando = "EXCLUI CARTOES RFID" Then
        xTempo = 50
    End If
    
    If Not AtualizaTabelaSolicitacaoAutomacao(pComando, pParametro) Then
        MsgBox "Não foi possível incluir SolicitaçãoFunçãoAutomação.", vbCritical, "Erro de Integridade!"
        Exit Function
    End If
    lNSU = SolicitacaoFuncaoAutomacao.NSU
    
    If Not AguardaSolicitacaoAutomacao Then
        Exit Function
    End If
    
    If Not AguardaSolicitAutoAprovado(xTempo) Then
        Exit Function
    Else
        If SolicitacaoFuncaoAutomacao.LocalizarNSU(lNSU) Then
            If SolicitacaoFuncaoAutomacao.HoraAprovacao <> "00:00:00" Then
                If SolicitacaoFuncaoAutomacao.TipoOperacao = "INCLUI CARTAO RFID" Then
                    CartaoAbastecimento.PosicaoRegistro = fRetiraString(SolicitacaoFuncaoAutomacao.Mensagem, 2)
                    If Len(CartaoAbastecimento.PosicaoRegistro) = 6 Then
                        If CartaoAbastecimento.Alterar(g_empresa, CartaoAbastecimento.NumeroCartao) Then
                            ComunicaAutomacaoCerradoBD = True
                            MsgBox "Cartão Incluído no IdentFID!", vbInformation + vbOKOnly, "Cartão Incluído!"
                        Else
                            MsgBox "Erro ao alterar o Registro do Cartão de Abastecimento!", vbOKOnly + vbCritical, "Cartão Não Alterado!"
                        End If
                    Else
                        MsgBox "Erro ao criticar retorno da Inclusão de Cartão de Abastecimento!" & vbCrLf & "Posicao do Registro=" & CartaoAbastecimento.PosicaoRegistro, vbCritical, "Cartão Não Incluído!"
                    End If
                ElseIf SolicitacaoFuncaoAutomacao.TipoOperacao = "AUTOMACAO LEITURA ENCERRANTE" Then
                    ComunicaAutomacaoCerradoBD = True
                ElseIf SolicitacaoFuncaoAutomacao.TipoOperacao = "AUTOMACAO ATIVADA" And fRetiraString(SolicitacaoFuncaoAutomacao.Mensagem, 1) = "OK" Then
                    ComunicaAutomacaoCerradoBD = True
                ElseIf SolicitacaoFuncaoAutomacao.TipoOperacao = "EXCLUI CARTOES RFID" Then
                    ComunicaAutomacaoCerradoBD = True
                Else
                    MsgBox "Operação desconhecida!" & vbCrLf & "TipoOperação=" & SolicitacaoFuncaoAutomacao.TipoOperacao, vbCritical, "Erro de Comando!"
                End If
            End If
        Else
            MsgBox "Não foi possível localizar SolicitaçãoFunçãoAutomação!" & vbCrLf & "NSU=" & lNSU, vbCritical, "Solicitação Função Inexistente!"
        End If
    End If
    
    
    
    Exit Function

FileError:
    Call CriaLogCupom("Erro ComunicaAutomacaoCerradoBD: Erro=" & Err.Number & " - " & Err.Description)
    MsgBox "Erro ao tentar comunicação com o programa AutoCerrado!", vbCritical, "Erro na Automação de Bomba!"
    Exit Function
End Function
'Function ComunicaPetromovelAutoBD() As Boolean
'    Dim xTempo As Integer
'    Dim xHoraInicial As Date
'    Dim xNSU As Long
'    Dim xMovSolicitacao As New CadastroDLL.cMovSolicitacaoFuncaoNFe
'
'    On Error GoTo FileError
'
'    ComunicaPetromovelAutoBD = False
'
'    xTempo = 10
'
'    xNSU = GravaSolicitacaoComunicacaoPetromovelAuto
'    If xNSU = 0 Then
'        'MsgBox "Não foi possível incluir registro de Solicitação de comunicação com PetromovelAuto!", vbCritical, "Erro de Integridade"
'        Exit Function
'    End If
'
'    If Not AguardaSolicitacaoPetromovelAuto(xTempo, xNSU) Then
'        Exit Function
'    End If
'
'    If xMovSolicitacao.LocalizarNSU(g_empresa, xNSU) Then
'        If xMovSolicitacao.HoraAnalise_MovSolicitacaoFuncaoNFe <> CDate("00:00:00") Then
'            If xMovSolicitacao.TipoOperacao_MovSolicitacaoFuncaoNFe = "PETROMOVEL_ATIVO" And xMovSolicitacao.Mensagem_MovSolicitacaoFuncaoNFe = "OK" Then
'                ComunicaPetromovelAutoBD = True
'            Else
'                Call CriaLogSGP("[ComunicaPetromovelAutoBD]", "Operação desconhecida!" & vbCrLf & "TipoOperação=" & xMovSolicitacao.TipoOperacao_MovSolicitacaoFuncaoNFe, "")
'                MsgBox "Operação desconhecida!" & vbCrLf & "TipoOperação=" & xMovSolicitacao.TipoOperacao_MovSolicitacaoFuncaoNFe, vbCritical, "Erro de Comando!"
'            End If
'        Else
'            Call xMovSolicitacao.DefineHoraCancelamentoHost(xNSU, Now, gVersaoSGP)
'        End If
'    Else
'        Call CriaLogSGP("[ComunicaPetromovelAutoBD]", "Não foi possível localizar a Solicitação de comunicação com o programa PetromovelAuto!", "")
'        MsgBox "Não foi possível localizar a Solicitação de comunicação com o programa PetromovelAuto!" & vbCrLf & "NSU=" & xNSU, vbCritical, "Solicitação Inexistente!"
'    End If
'
'    Exit Function
'
'FileError:
'    Call CriaLogSGP("[ComunicaPetromovelAutoBD]", "Erro ao tentar comunicação com o programa PetromovelAuto - Erro=" & Err.Number & " - " & Err.Description, "")
'    'MsgBox "Erro ao tentar comunicação com o programa PetromovelAuto!", vbCritical, "Erro na emissão de NFCe!"
'    Exit Function
'End Function
Function ComunicaAutomacaoIonics(ByVal pComando As String, ByVal pParametro As String) As Boolean
    Dim xArquivoTmp As String
    Dim xArquivoPedido As String
    Dim xArquivoResp As String
    Dim xComputadorAutomacao As String
    Dim xPastaAutomacao As String
    Dim xHoraInicial As Date
    Dim xComando As String
    Dim xRetorno As String
    Dim xParametro As String
    Dim xTempo As Integer
    'Dim xNumeroEmailInicial As String

    On Error GoTo FileError

    ComunicaAutomacaoIonics = False

    'Grava tipo de email
    'xNumeroEmailInicial = gNumeroEmailInicial
    'Call WriteINI("EMAIL", "Numero do Email", xNumeroEmailInicial, lNomeArquivoAutomacaoIni)
    'If gNumeroEmailInicial > 0 Then
    '    Call WriteINI("EMAIL", "Concluido", "NAO", lNomeArquivoAutomacaoIni)
    'Else
    '    Call WriteINI("EMAIL", "Concluido", "SIM", lNomeArquivoAutomacaoIni)
    'End If

    'Pega o NOME do computador que tem Ligado Fisicamente o Equipamento de Automação
    xComputadorAutomacao = ReadINI("LOCALIZACAO", "Computador com automacao", lNomeArquivoAutomacaoIni)
    xPastaAutomacao = ReadINI("LOCALIZACAO", "Pasta da automacao", lNomeArquivoAutomacaoIni)
    If xPastaAutomacao = "" Then
        xPastaAutomacao = "Automacao"
    End If
    'MsgBox "xComputadorAutomacao=" & xComputadorAutomacao
    If xComputadorAutomacao = "" Then
        xComputadorAutomacao = GetIPHostName()
    End If
    'MsgBox "xComputadorAutomacao=" & xComputadorAutomacao

    'Monta nome do Computador + Diretório + Arquivo do Pedido de comunicação
    'Ex: \\Servidor\Automacao\Pedido_ddmmyyyy_HHmmss.TMP
    'xArquivoTmp = "\\" & xComputadorAutomacao & "\Automacao\Pedido_" & Format(Date, "ddmmyyyy") & "_" & Format(Time, "HHmmss") & ".TMP"
    xArquivoTmp = "\\" & xComputadorAutomacao & "\" & xPastaAutomacao & "\Solicitacao.TMP"
    xArquivoPedido = Mid(xArquivoTmp, 1, Len(xArquivoTmp) - 3) & "TXT"

    'Cria o arquivo .TMP de comunicação
    Set gArquivoTXT = gArqTxt.CreateTextFile(xArquivoTmp)
    gArquivoTXT.WriteLine ("[PEDIDO AUTOMACAO]")
    gArquivoTXT.WriteLine ("Comando=" & pComando)
    gArquivoTXT.WriteLine ("Origem=" & GetIPHostName())
    gArquivoTXT.WriteLine ("Parametro=" & pParametro)
    gArquivoTXT.Close

    'Aguarda 1 Segundo
    'Para antivirus nao bloquear a renomeação do arquivo
    xTempo = 1
    xHoraInicial = Time
    Do Until DateDiff("s", xHoraInicial, Time) >= xTempo
        If gArqTxt.FileExists(xArquivoResp) Then
            Exit Do
        End If
        DoEvents
    Loop

    'Renomeia arquivo .TMP para .AUT
    If gArqTxt.FileExists(xArquivoTmp) Then
        gArqTxt.MoveFile (xArquivoTmp), (xArquivoPedido)
    End If

    'Monta nome do Arquivo de Retorno
    'xArquivoResp = "C:\Automacao\Retorno_" & Mid(xArquivoPedido, Len(xArquivoPedido) - 19 + 1, 19)
    'xArquivoResp = "C:\Automacao\Resposta.TXT"
    xArquivoResp = "\\" & xComputadorAutomacao & "\" & xPastaAutomacao & "\Resposta.TXT"

    xTempo = 2
    If pComando = "AUTOMACAO LEITURA ENCERRANTE" Then
        xTempo = lQtdBomba * xTempo
    End If

    'Aguarda até (2 * lQtdBomba) Segundos para o retorno
    xHoraInicial = Time
    Do Until DateDiff("s", xHoraInicial, Time) >= xTempo
        If gArqTxt.FileExists(xArquivoResp) Then
            Exit Do
        End If
        DoEvents
    Loop

    'Verifica se o Retorno existe
    If gArqTxt.FileExists(xArquivoResp) Then
        'Existindo lê o retorno
        xComando = ReadINI("RETORNO AUTOMACAO", "Comando", xArquivoResp)
        xRetorno = ReadINI("RETORNO AUTOMACAO", "Retorno", xArquivoResp)
        xParametro = ReadINI("RETORNO AUTOMACAO", "Parametro", xArquivoResp)
        'Deleta o arquivo de retorno retorno
        gArqTxt.DeleteFile (xArquivoResp)
        If xRetorno = "OK" Then
            ComunicaAutomacaoIonics = True
        End If
    Else
        'MsgBox "arquivo nao encontrado=" & xArquivoResp
        'Deleta o Arquivo de pedido
        'Pois fica sub-entendido que o mesmo ainda existe
        gArqTxt.DeleteFile (xArquivoPedido)
    End If

    Exit Function

FileError:
    Call CriaLogCupom("Erro ComunicaAutomacaoIonics: Erro=" & Err.Number & " - " & Err.Description)
    Call CriaLogCupom("     Arquivo xArquivoTmp: " & xArquivoTmp)
    Call CriaLogCupom("     Arquivo xArquivoPedido: " & xArquivoPedido)
    MsgBox "Erro ao tentar comunicação com o programa AutoCerradoEZ!", vbCritical, "Erro na Automação de Bomba!"
    Exit Function
End Function
Private Function CriaAberturaCaixa(ByVal pTipoMovimento As Integer, ByVal pPeriodo As Integer) As Boolean

    On Error GoTo FileError
    
    CriaAberturaCaixa = False
    AberturaCaixa.Empresa = g_empresa
    AberturaCaixa.DataAbertura = Format(CDate(Date), "dd/mm/yyyy")
    AberturaCaixa.TipoCaixa = "NF"
    AberturaCaixa.Periodo = pPeriodo
    AberturaCaixa.NumeroIlha = lIlha
    AberturaCaixa.CodigoFuncionario = l_codigo_funcionario
    AberturaCaixa.HoraAbertura = Format(Time, "hh:mm:ss")
    AberturaCaixa.DataFechamento = "00:00:00"
    AberturaCaixa.HoraFechamento = "00:00:00"
    AberturaCaixa.TipoMovimento = pTipoMovimento
    AberturaCaixa.FechadoPeloNivel = 0
    AberturaCaixa.RecebidoPeloFinanceiro = False
    AberturaCaixa.DataConferencia = "00:00:00"
    AberturaCaixa.ConferidopeloNivel = 0
    If AberturaCaixa.Incluir = False Then
        MsgBox "Não foi possível abrir o caixa!", vbCritical, "Erro de Integridade!"
    Else
        CriaAberturaCaixa = True
    End If
    Exit Function

FileError:
    Call CriaLogCupom("Erro CriaAberturaCaixa: Erro=" & Err.Number & " - " & Err.Description)
    MsgBox "Erro ao criar abertura de caixa!", vbCritical, "Erro desconhecido!"
    Exit Function
End Function
'INICIO NEW 07/04
Private Sub SelecionaVeiculoCliente(ByVal pCodigoCliente As Long)
    Dim xString As String
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    
    If VeiculoCliente.ClienteTemVeiculo(pCodigoCliente) Then
        lSQL = "SELECT [Codigo do Veiculo], Nome, Cor, Ano, [Placa Letra], [Placa Numero] FROM VeiculoCliente WHERE [Codigo do Cliente] = " & pCodigoCliente & " ORDER BY Nome"
        Set rs = Conectar.RsConexao(lSQL)
        i = 0
        xString = ""
        g_string = ""
        With rs
            If .RecordCount > 0 Then
                .MoveFirst
                Do Until .EOF
                    i = i + 1
                    xString = xString & rs(0).Value & "|@|"
                    xString = xString & Trim(rs(1).Value) & ", " & Trim(rs(2).Value) & ", " & rs(3).Value & ", " & rs(4).Value & "-" & rs(5).Value & "|@|"
                    .MoveNext
                Loop
            End If
            .Close
        End With
        Set rs = Nothing
        If i > 0 Then
            xString = "Selecione o Veículo Desejado|@|" & i & "|@|" & xString
            Do Until (Len(g_string) = 3 Or Len(g_string) = 4)
                g_string = xString
                opcaoGeral.Show 1
                If Len(g_string) > 0 Then
                    lCodigoVeiculo = RetiraGString(1)
                    Exit Do
                End If
            Loop
            If Not VeiculoCliente.LocalizarCodigo(pCodigoCliente, lCodigoVeiculo) Then
                MsgBox "Veículo de cliente não foi localizado!", vbInformation, "Dados Incompleto!"
            End If
        End If
    Else
        Exit Sub
    End If
End Sub
'FIM NEW 07/04

Private Sub DarumaBuscaRetorno()
    lAck = 0
    lSt1 = 0
    lSt2 = 0
    'BemaRetorno = Daruma_FI_RetornoImpressora(lAck, lSt1, lSt2)
    'lErroExtendido = Space(4)
    'BemaRetorno = Daruma_FI_RetornaErroExtendido(lErroExtendido)
End Sub
Private Sub DefinePortaEcf()
    Dim xPortaEcf As String

    xPortaEcf = ReadINI("CUPOM FISCAL", "Porta ECF", gArquivoIni)
    If xPortaEcf = "" Then
        xPortaEcf = "COM1"
    End If
    
    Call WriteINI("Sistema", "Porta", xPortaEcf, "c:\windows\system32\bemafi32.ini")
End Sub
Private Sub DescarregaAbastecimento()
    Dim i As Integer
    Dim xCupomImpresso As Integer
    Dim xTotalBruto As Currency
    Dim xTotalLiquido As Currency
    Dim xSQL As String
    
    On Error GoTo FileError
    
    xCupomImpresso = 0
    
    xSQL = ""
    xSQL = xSQL & "SELECT Data, Hora, Bico"
    xSQL = xSQL & "  FROM Movimento_Abastecimento"
    xSQL = xSQL & " WHERE Empresa = " & g_empresa
    xSQL = xSQL & "   AND Bico = " & Val(cboBico.Text)
    xSQL = xSQL & "   AND Acerto = " & preparaBooleano(False)
    xSQL = xSQL & " ORDER BY Data, Hora"
    rstAbastecimento.Open xSQL, cnnSGP, adOpenForwardOnly, adLockReadOnly
    If rstAbastecimento.RecordCount > 0 Then
        Do Until rstAbastecimento.EOF
            If MovimentoAbastecimento.LocalizarCodigo(g_empresa, rstAbastecimento("Data").Value, rstAbastecimento("Hora").Value, rstAbastecimento("Bico").Value) Then
                'Caso foi escolhido para ter desconto
                'Pega o valor do "Preço Médio" no cadastro de combustível
                'E calcula a diferença para jogar como desconto
                
                If chkDesconto.Value = 1 Then
                    If Bomba.LocalizarCodigo(g_empresa, MovimentoAbastecimento.Bico) Then
                        If Combustivel.LocalizarCodigo(g_empresa, Bomba.TipoCombustivel) Then
                            If Combustivel.PrecoMedio < Bomba.PrecoVenda And Combustivel.PrecoMedio > Bomba.PrecoCusto Then
                                xTotalBruto = MovimentoAbastecimento.ValorTotal
                                xTotalLiquido = Combustivel.PrecoMedio * MovimentoAbastecimento.Quantidade
                                lDescontoEspecial = lDescontoEspecial + (xTotalBruto - xTotalLiquido)
                                'Call CriaLogCupom("Cupom Fiscal (Teste Desconto): xTotalLiquido=" & xTotalLiquido & " - xTotalBruto=" & xTotalBruto & " - lDescontoEspecial=" & lDescontoEspecial)
                            Else
                                Call CriaLogCupom("Cupom Fiscal (Combustível Inexistente): Combustivel.PrecoMedio=" & Combustivel.PrecoMedio & " - Bomba.PrecoVenda=" & Bomba.PrecoVenda)
                            End If
                        Else
                            MsgBox "Não foi possível localizar o combustível " & Bomba.TipoCombustivel & ".", vbInformation, "Erro de Integridade!"
                        End If
                    Else
                        MsgBox "Não foi possível localizar o bico " & MovimentoAbastecimento.Bico & ".", vbInformation, "Erro de Integridade!"
                    End If
                End If
                txt_produto.Text = MovimentoAbastecimento.CodigoProduto
                txt_quantidade.Text = MovimentoAbastecimento.Quantidade
                txt_valor_total.Text = MovimentoAbastecimento.ValorTotal
                txt_produto_LostFocus
                txt_quantidade.Text = Format(MovimentoAbastecimento.Quantidade, "###,##0.00")
                txt_valor_total.Text = Format(MovimentoAbastecimento.ValorTotal, "###,##0.00")
                lOrigemAutomacao = True
                lAutomacaoFlagVendaAutomatica = True
                lAutoBico = MovimentoAbastecimento.Bico
                lAutoQuantidade = MovimentoAbastecimento.Quantidade
                lAutoValorTotal = MovimentoAbastecimento.ValorTotal
                lAutoHora = MovimentoAbastecimento.Hora
                lAutomacaoBicoEmAcerto = MovimentoAbastecimento.Bico
                lAutomacaoDataEmAcerto = MovimentoAbastecimento.Data
                lAutomacaoHoraEmAcerto = MovimentoAbastecimento.Hora
                lAutomacaoTempoAbastecimentoEmAcerto = "" & MovimentoAbastecimento.TempoAbastecimento
                lAutomacaoFormaImpressao = "DESC."
                Call CriaLogCupom("Cupom Fiscal: Descarregando bico=" & lAutoBico & " - Quantidade=" & lAutoQuantidade & " - Valor Total=" & lAutoValorTotal & " - ECF=" & lNumeroCupom & " - Hora do Abastecimento=" & lAutoHora)
                GravaItem
                MovimentoAbastecimento.Acerto = True
                lAutomacaoBicoEmAcerto = 0
                lAutomacaoDataEmAcerto = 0
                lAutomacaoHoraEmAcerto = 0
                lAutomacaoTempoAbastecimentoEmAcerto = ""
                lAutomacaoFormaImpressao = ""
                MovimentoAbastecimento.NumeroCupom = lNumeroCupom
                MovimentoAbastecimento.CodigoECF = lNumeroPDV
                MovimentoAbastecimento.DocumentoGerado = "CP"
                If Not MovimentoAbastecimento.Alterar(g_empresa, rstAbastecimento("Data").Value, rstAbastecimento("Hora").Value, rstAbastecimento("Bico").Value) Then
                    MsgBox "Não foi possível alterar o abastecimento!", vbInformation, "Erro de Integridade!"
                End If
                xCupomImpresso = xCupomImpresso + 1
                If xCupomImpresso = Val(txtQuantidadeDescarregamento.Text) Then
                    Exit Do
                End If
            Else
                MsgBox "Não foi possível localizar o abastecimento!", vbInformation, "Erro de Integridade!"
            End If
            rstAbastecimento.MoveNext
        Loop
    End If
    rstAbastecimento.Close
    
    lOrigemAutomacao = False
    lAutomacaoFlagVendaAutomatica = False
    frm_fila_bico.Enabled = False
    frm_fila_bico.Visible = False
    cmd_bico(0).SetFocus
    Exit Sub
    
FileError:
    Call CriaLogCupom("Cupom Fiscal: ERRO DescarregaAbastecimento - Bico=" & lAutoBico & " - Quantidade=" & lAutoQuantidade & " - Valor Total=" & lAutoValorTotal & " - ECF=" & lNumeroCupom & " - Hora do Abastecimento=" & lAutoHora)
    MsgBox "Erro na Rotina: DescarregaAbastecimento", vbInformation, "Erro Desconhecido"
    Exit Sub
End Sub
Private Sub DespreparaDadosAdicionaisFechamento()
    cmdIniciaProcessoFinalizacaoNFCe.Visible = False
    lbl_numero_cheque.Visible = False
    txt_numero_cheque.Visible = False
    lbl_telefone.Visible = False
    txt_telefone.Visible = False
    lbl_valor_recebido.Left = 3180
    txt_valor_recebido.Left = 3180
    lbl_valor_troco1.Left = 4620
    lbl_valor_troco.Left = 4620
    cmd_cancelar2.Left = 4000
    cmdIniciaProcessoNFCe.Left = 4920
    cmd_cancelar2.Top = 3280
    cmdIniciaProcessoNFCe.Top = 3280
    
    cmdDescontoAppPostoAki.Top = 8500 '6680
    cmdDescontoAppPostoAki.Left = 100
    cmdCartaoFidelidadeDesconto.Top = 8500 '6680
    cmdCartaoFidelidadeDesconto.Left = 2000
    cmd_DescontoCorreio.Top = 8500 '6680
    cmd_DescontoCorreio.Left = 3900
    
    btnDescontoCartaoEspecial.Top = 8500 '6680
    btnDescontoCartaoEspecial.Left = 5800
    btnDescontoCartaoEspecial.Enabled = False
    
    cmdDescontoAppPostoAki.Enabled = False
    'cmdCartaoFidelidadeDesconto.Enabled = False - ALEX
    cmd_DescontoCorreio.Enabled = False
    
    If ConfiguracaoDiversa.LocalizarCodigo(1, "PERCENTUAL CARTAO DESCONTO CORREIOS") Then
        If ConfiguracaoDiversa.Valor > 0 Then
            cmd_DescontoCorreio.Enabled = True
        End If
    End If
    If ConfiguracaoDiversa.LocalizarCodigo(1, "DESCONTO APP POSTOAKI") Then
        If ConfiguracaoDiversa.Verdadeiro = True Then
            cmdDescontoAppPostoAki.Enabled = True
        End If
    End If
    If cbo_forma_pagamento.ListIndex >= 0 Then
        If cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 16 Then 'POS
            If ConfiguracaoDiversa.LocalizarCodigo(1, "PERCENTUAL DESCONTO CARTAO ESPECIAL") Then
                If ConfiguracaoDiversa.Valor > 0 Then
                    btnDescontoCartaoEspecial.Enabled = True
                End If
            End If
        ElseIf lDescontoCartaoEspecial Then
            txt_valor_desconto.Text = Format(0, "###,##0.00")
            lValorDescontoConcedido = 0
            
            If lRestringeDescontoNivelUsuario Then
                txt_valor_desconto.Enabled = False
            Else
                txt_valor_desconto.Enabled = True
            End If
            
            
            lbl_valor_compra.Caption = Format(lTotalCupom - fValidaValor(txt_valor_desconto.Text), "###,##0.00")
            txt_valor_recebido.Text = Format(lTotalCupom - fValidaValor(txt_valor_desconto.Text), "###,##0.00")

            lDescontoCartaoEspecial = False
        End If
    Else
        lDescontoCartaoEspecial = False
    End If
    
    frm_fechamento_cupom.Top = 50 '100
    frm_fechamento_cupom.Left = 30 '70
    frm_fechamento_cupom.Height = 9500 '7480
    frm_fechamento_cupom.Width = 7850 '7460
End Sub
Private Function ExisteCaixaIndividualAberto(ByVal pData As Date) As Boolean
    Dim xPeriodo As Integer
    Dim xString As String
    
    ExisteCaixaIndividualAberto = False
    'MsgBox "Cod func:" & l_codigo_funcionario & " Data:" & pData, vbCritical, "Teste"
    'If Not AberturaCaixa.LocalizarUltAbertoDataFunc(g_empresa, pData, "NF", 1, 2, l_codigo_funcionario) Then
    If Not AberturaCaixa.LocalizarUltAbertoDataFunc(g_empresa, pData, "NF", 1, lTipoMovimento, l_codigo_funcionario) Then
        'Informa período do caixa a ser aberto
        If lTipoMovimento = 3 Then
            xPeriodo = lPeriodo
        Else
            xString = InputBox("Informe número do período que deseja abrir.", "Período à Abrir!", "")
            If Val(xString) = 0 Or Val(xString) > 4 Then
                MsgBox "O período informado não é válido.", vbOKOnly + vbInformation, "Período Inválido!"
                txt_senha_ponto.Text = ""
                txt_senha_ponto.SetFocus
                mnuSenha_Click
                Exit Function
            End If
            xPeriodo = Val(xString)
        End If
        'verifica se existe caixa no período informado
        'If AberturaCaixa.LocalizarCodigo(g_empresa, pData, "NF", xPeriodo, 1, l_codigo_funcionario, 2) Then
        If AberturaCaixa.LocalizarCodigo(g_empresa, pData, "NF", xPeriodo, 1, l_codigo_funcionario, lTipoMovimento) Then
            MsgBox "Já existe um caixa no período informado.", vbOKOnly + vbInformation, "Período Inválido!"
            Exit Function
        End If
        If Not CartaoAbastecimento.LocalizarCodigoFuncionario(g_empresa, l_codigo_funcionario) Then
            MsgBox "Este funcionário não tem cartão de abastecimento vinculado.", vbOKOnly + vbCritical, "Erro de Integridade!"
            Exit Function
        End If
        If (MsgBox("O Caixa não encontra-se aberto!" & Chr(10) & "Deseja abrir agora?", vbQuestion + vbYesNo + vbDefaultButton1, "Abertura de Caixa!")) = vbYes Then
            'TIPO DE MOVIMENTO = 2 CX PISTA
            'PERIODO = 1
            'If lMarcaAutomacao = "COMPANY" And (Not UCase(g_nome_empresa) Like "*VERA CRUZ*") Then 'Or lMarcaAutomacao = "HOROUSTECH" Then
            'If lMarcaAutomacao = "COMPANY" And (UCase(g_nome_empresa) Like "*POSTO MT LTDA*" Or UCase(g_nome_empresa) Like "*BRITO BARROS LTDA*") Then
            If lMarcaAutomacao = "COMPANY" And (UCase(g_nome_empresa) Like "*BRITO BARROS LTDA*") Then
                If ComunicaAutomacaoCerradoBD("INCLUI CARTAO RFID", CartaoAbastecimento.NumeroCartao & "|@|") Then
                    If CriaAberturaCaixa(2, xPeriodo) Then
                        ExisteCaixaIndividualAberto = True
                    End If
                Else
                    If lMarcaAutomacao = "COMPANY" Then
                        MsgBox "Não foi possível comunicar com o progama AutoCerradoCompany.", vbCritical, "Erro de Automação!"
                    ElseIf lMarcaAutomacao = "HOROUSTECH" Then
                        MsgBox "Não foi possível comunicar com o progama AutoCerradoHorousTech.", vbCritical, "Erro de Automação!"
                    End If
                End If
            ElseIf lMarcaAutomacao = "EZTECH" Or lMarcaAutomacao = "HOROUSTECH" Or lMarcaAutomacao = "COMPANY" Then
                'QUANDO FOR DEBUGAR AUTOMACAO PULAR A LINHA ABAIXO, PARA ENTRAR DENTRO DO TESTE
                If ComunicaAutomacaoCerradoBD("AUTOMACAO ATIVADA", "" & "|@|") Then
                    'If CriaAberturaCaixa(2, xPeriodo) Then
                    If CriaAberturaCaixa(lTipoMovimento, xPeriodo) Then
                        ExisteCaixaIndividualAberto = True
                    End If
                Else
                    MsgBox "Não foi possível comunicar com o progama AutoCerradoEZ.", vbCritical, "Erro de Automação!"
                End If
            End If
        Else
            Exit Function
        End If
    Else
        If AberturaCaixa.DataFechamento = "00:00:00" Then
            ExisteCaixaIndividualAberto = True
        Else
            MsgBox "Este funcionário está com o caixa de hoje fechado." & vbCrLf & "Data do Fechamento=" & Format(AberturaCaixa.DataFechamento, "dd/MM/yyyy") & vbCrLf & "Hora do Fechamento=" & Format(AberturaCaixa.HoraFechamento, "HH:mm:ss"), vbOKOnly + vbInformation, "Operação Negada!"
            mnuFuncao.Enabled = True
            mnuTCS.Enabled = False
            mnuFuncaoADM.Enabled = False
            mnuCancelaCartao.Enabled = False
            mnuLancamentoEncerrante.Enabled = False
            mnuMudaProximoTurno.Enabled = False
            mnuPontoFuncionario.Enabled = False
           ' mnuReducaoZ.Enabled = False
            mnuFechamentoCaixa.Enabled = True
            Exit Function
        End If
    End If
End Function
Function ExcluiMovimentoCaixaNFCe(ByVal pTipoLancamentoPadrao As String, ByVal pDataEmissao As Date, ByVal pPeriodo As Integer, ByVal pValorTotal As Currency, ByVal pTipoSubEstoque As Integer) As Boolean
    Dim xComplemento As String
    Dim xValor As Currency
    
    On Error GoTo trata_erro
    
    ExcluiMovimentoCaixaNFCe = False
    xValor = 0
    xComplemento = ""
    If Not IntegracaoCaixa.LocalizarNome(g_empresa, pTipoLancamentoPadrao) Then
        MsgBox "Não existe a integração=" & pTipoLancamentoPadrao & ".", vbInformation, "Registro Inexistente"
        Call GravaAuditoria(1, Me.name, 25, "Não será integrado no caixa o extorno de:" & pTipoLancamentoPadrao)
        Exit Function
    End If
    If pTipoLancamentoPadrao = "VENDA DE LUBRIFICANTES" Then
        xComplemento = "LUBRIFICANTES Per:" & pPeriodo & " Ilha:" & lIlha & " S.Est:" & pTipoSubEstoque & " T.Mov:" & lTipoMovimento
        Call CriaLogCupom("Teste ExcluiMovimentoCaixa: xComplemento:" & xComplemento)
        If MovCaixaPista.LocalizarRegistroEspecialUsu(g_empresa, pDataEmissao, pPeriodo, lIlha, xComplemento, IntegracaoCaixa.ContaCredito, "C", g_usuario) Then
            xValor = MovCaixaPista.Valor - pValorTotal
            If xValor = 0 Then
                If MovCaixaPista.Excluir(g_empresa, pDataEmissao, MovCaixaPista.NumeroMovimento) Then
                    ExcluiMovimentoCaixaNFCe = True
                Else
                    MsgBox "Não foi possível excluir registro especial de produto no caixa.", vbCritical, "Erro de Integridade!"
                    Call GravaAuditoria(1, Me.name, 25, "Não foi possível excluir registro especial de produto no caixa.")
                    Call GravaAuditoria(1, Me.name, 25, "Data:" & pDataEmissao & " Numero Mov:" & MovCaixaPista.NumeroMovimento)
                End If
            Else
                MovCaixaPista.Valor = xValor
                MovCaixaPista.DataAlteracao = Format(Now, "dd/mm/yyyy")
                MovCaixaPista.HoraAlteracao = Format(Now, "HH:mm:ss")
                If MovCaixaPista.Alterar(g_empresa, pDataEmissao, MovCaixaPista.NumeroMovimento) Then
                    ExcluiMovimentoCaixaNFCe = True
                Else
                    MsgBox "Não foi possível alterar registro especial de produto no caixa.", vbCritical, "Erro de Integridade!"
                    Call GravaAuditoria(1, Me.name, 25, "Não foi possível alterar registro especial de produto no caixa.")
                    Call GravaAuditoria(1, Me.name, 25, "Data:" & pDataEmissao & " Numero Mov:" & MovCaixaPista.NumeroMovimento)
                End If
            End If
        Else
            MsgBox "Não foi possível localiar registro especial no caixa.", vbCritical, "Erro de Integridade!"
            Call GravaAuditoria(1, Me.name, 25, "Não foi possível localizar registro especial de produto no caixa.")
            Call GravaAuditoria(1, Me.name, 25, "Data:" & pDataEmissao & " Per:" & pPeriodo & " Ilha:" & lIlha & " Credito")
            Call GravaAuditoria(1, Me.name, 25, "Comp:" & xComplemento & " Conta:" & IntegracaoCaixa.ContaCredito)
            Exit Function
        End If
    End If
    
    Exit Function

trata_erro:
    Call CriaLogCupom("Erro ExcluiMovimentoCaixa: Erro=" & Err.Number & " - " & Err.Description)
    Call GravaAuditoria(1, Me.name, 25, "ExcluiMovimentoCaixaNFCe: Erro inesperado...")
End Function
Function ExcluiMovimentoCaixa(ByVal pTipoLancamentoPadrao As String) As Boolean
    Dim xComplemento As String
    Dim xValor As Currency

    On Error GoTo trata_erro
    
    ExcluiMovimentoCaixa = False
    xValor = 0
    xComplemento = ""
    If Not IntegracaoCaixa.LocalizarNome(g_empresa, pTipoLancamentoPadrao) Then
        MsgBox "Não existe a integração=" & pTipoLancamentoPadrao & ".", vbInformation, "Registro Inexistente"
        Call GravaAuditoria(1, Me.name, 25, "Não será integrado no caixa o extorno de:" & pTipoLancamentoPadrao)
        Exit Function
    End If
    If pTipoLancamentoPadrao = "VENDA DE LUBRIFICANTES" Then
        xComplemento = "LUBRIFICANTES Per:" & MovCupomFiscal.Periodo & " Ilha:" & lIlha & " S.Est:" & MovCupomFiscal.TipoSubEstoque & " T.Mov:" & lTipoMovimento
        Call CriaLogCupom("Teste ExcluiMovimentoCaixa: xComplemento:" & xComplemento)
        If MovCaixaPista.LocalizarRegistroEspecial(g_empresa, MovCupomFiscal.Data, Val(MovCupomFiscal.Periodo), lIlha, xComplemento, IntegracaoCaixa.ContaCredito, "C") Then
            xValor = MovCaixaPista.Valor - MovCupomFiscal.ValorTotal
            If xValor = 0 Then
                If MovCaixaPista.Excluir(g_empresa, MovCupomFiscal.Data, MovCaixaPista.NumeroMovimento) Then
                    ExcluiMovimentoCaixa = True
                Else
                    MsgBox "Não foi possível excluir registro especial de produto no caixa.", vbCritical, "Erro de Integridade!"
                    Call GravaAuditoria(1, Me.name, 25, "Não foi possível excluir registro especial de produto no caixa.")
                    Call GravaAuditoria(1, Me.name, 25, "Data:" & MovCupomFiscal.Data & " Numero Mov:" & MovCaixaPista.NumeroMovimento)
                End If
            Else
                MovCaixaPista.Valor = xValor
                MovCaixaPista.DataAlteracao = Format(Now, "dd/mm/yyyy")
                MovCaixaPista.HoraAlteracao = Format(Now, "HH:mm:ss")
                If MovCaixaPista.Alterar(g_empresa, MovCupomFiscal.Data, MovCaixaPista.NumeroMovimento) Then
                    ExcluiMovimentoCaixa = True
                Else
                    MsgBox "Não foi possível alterar registro especial de produto no caixa.", vbCritical, "Erro de Integridade!"
                    Call GravaAuditoria(1, Me.name, 25, "Não foi possível alterar registro especial de produto no caixa.")
                    Call GravaAuditoria(1, Me.name, 25, "Data:" & MovCupomFiscal.Data & " Numero Mov:" & MovCaixaPista.NumeroMovimento)
                End If
            End If
        Else
            MsgBox "Não foi possível localiar registro especial no caixa.", vbCritical, "Erro de Integridade!"
            Call GravaAuditoria(1, Me.name, 25, "Não foi possível localizar registro especial de produto no caixa.")
            Call GravaAuditoria(1, Me.name, 25, "Data:" & MovCupomFiscal.Data & " Per:" & MovCupomFiscal.Periodo & " Ilha:" & lIlha & " Credito")
            Call GravaAuditoria(1, Me.name, 25, "Comp:" & xComplemento & " Conta:" & IntegracaoCaixa.ContaCredito)
            Exit Function
        End If
    ElseIf pTipoLancamentoPadrao = "NOTA ABASTECIMENTO" Then
        If Cliente.LocalizarCodigo(MovCupomFiscal.CodigoCliente) Then
            xComplemento = Cliente.RazaoSocial
        End If
        If MovCaixaPista.LocalizarRegistroEspecialDoc(g_empresa, MovCupomFiscal.Data, Val(MovCupomFiscal.Periodo), lIlha, xComplemento, Format(MovCupomFiscal.NumeroCupom, "#######0") & Format(MovCupomFiscal.Ordem, "00"), IntegracaoCaixa.ContaDebito, "D") Then
            If MovCaixaPista.Excluir(g_empresa, MovCupomFiscal.Data, MovCaixaPista.NumeroMovimento) Then
                ExcluiMovimentoCaixa = True
            Else
                MsgBox "Não foi possível excluir registro especial de Notas no caixa.", vbCritical, "Erro de Integridade!"
                Call GravaAuditoria(1, Me.name, 25, "Não foi possível excluir registro especial de Notas no caixa.")
                Call GravaAuditoria(1, Me.name, 25, "Data:" & MovCupomFiscal.Data & " Numero Mov:" & MovCaixaPista.NumeroMovimento)
            End If
        Else
            'MsgBox "Não foi possível localiar registro especial no caixa.", vbCritical, "Erro de Integridade!"
            Call GravaAuditoria(1, Me.name, 25, "Não foi possível localizar registro especial de Notas no caixa.")
            Call GravaAuditoria(1, Me.name, 25, "Data:" & MovCupomFiscal.Data & " Per:" & MovCupomFiscal.Periodo & " Ilha:" & lIlha & " Debito. N.Doc:" & Format(MovCupomFiscal.NumeroCupom, "#######0") & Format(MovCupomFiscal.Ordem, "00"))
            Call GravaAuditoria(1, Me.name, 25, "Comp:" & xComplemento & " Conta:" & IntegracaoCaixa.ContaDebito)
            Exit Function
        End If
    ElseIf pTipoLancamentoPadrao = "NOTA ABASTECIMENTO DESCONTO" Or pTipoLancamentoPadrao = "NOTA ABASTECIMENTO ACRESCIMO" Then
        If Cliente.LocalizarCodigo(MovCupomFiscal.CodigoCliente) Then
            xComplemento = Cliente.RazaoSocial
        End If
        If MovCaixaPista.LocalizarRegistroEspecialDoc(g_empresa, MovCupomFiscal.Data, Val(MovCupomFiscal.Periodo), lIlha, xComplemento, Format(MovCupomFiscal.NumeroCupom, "#######0") & Format(MovCupomFiscal.Ordem, "00"), IntegracaoCaixa.ContaDebito, "D") Then
            If MovCaixaPista.Excluir(g_empresa, MovCupomFiscal.Data, MovCaixaPista.NumeroMovimento) Then
                ExcluiMovimentoCaixa = True
            Else
                MsgBox "Não foi possível excluir registro especial de Notas Desc. no caixa.", vbCritical, "Erro de Integridade!"
                Call GravaAuditoria(1, Me.name, 25, "Não foi possível excluir registro especial de Notas Desc. no caixa.")
                Call GravaAuditoria(1, Me.name, 25, "Data:" & MovCupomFiscal.Data & " Numero Mov:" & MovCaixaPista.NumeroMovimento)
            End If
        Else
            'MsgBox "Não foi possível localiar registro especial no caixa.", vbCritical, "Erro de Integridade!"
            Call GravaAuditoria(1, Me.name, 25, "Não foi possível localizar registro especial de Notas Desc.no caixa.")
            Call GravaAuditoria(1, Me.name, 25, "Data:" & MovCupomFiscal.Data & " Per:" & MovCupomFiscal.Periodo & " Ilha:" & lIlha & " Debito. N.Doc:" & Format(MovCupomFiscal.NumeroCupom, "#######0") & Format(MovCupomFiscal.Ordem, "00"))
            Call GravaAuditoria(1, Me.name, 25, "Comp:" & xComplemento & " Conta:" & IntegracaoCaixa.ContaDebito)
            Exit Function
        End If
    End If
    Exit Function

trata_erro:
    Call CriaLogCupom("Erro ExcluiMovimentoCaixa: Erro=" & Err.Number & " - " & Err.Description)
    Call GravaAuditoria(1, Me.name, 25, "ExcluiMovimentoCaixa: Erro inesperado...")
End Function
Private Sub ExcluiNotaAbastecimento()
    Dim xTipoMovimento As Integer
    Dim xStringIntegracao As String
    Dim i As Integer

On Error GoTo trata_erro
    
    xTipoMovimento = 2
    If lLoja Then
        xTipoMovimento = 1
    End If

    For i = 1 To 3
        If i = 1 Then
            xStringIntegracao = "NOTA ABASTECIMENTO"
        ElseIf i = 2 Then
            xStringIntegracao = "NOTA ABASTECIMENTO DESCONTO"
        ElseIf i = 3 Then
            xStringIntegracao = "NOTA ABASTECIMENTO ACRESCIMO"
        End If
        If Not IntegracaoCaixa.LocalizarNome(g_empresa, xStringIntegracao) Then
            MsgBox "Não será possível integrar com o caixa!", vbCritical, "Erro de Integridade!"
            Call GravaAuditoria(1, Me.name, 25, "Não será integrado no caixa o extorno de:" & xStringIntegracao)
            Exit For
        End If
        If ExcluiMovimentoCaixa(xStringIntegracao) Then
            If i = 1 Then
                If MovNotaAbastecimento.LocalizarCodigo(g_empresa, MovCupomFiscal.CodigoCliente, MovCupomFiscal.Data, Format(MovCupomFiscal.NumeroCupom, "#######0") & Format(MovCupomFiscal.Ordem, "00"), MovCupomFiscal.Ordem, MovCupomFiscal.CodigoProduto, MovCupomFiscal.Periodo) Then
                    If MovNotaAbastecimento.Excluir(g_empresa, MovCupomFiscal.CodigoCliente, MovCupomFiscal.Data, Format(MovCupomFiscal.NumeroCupom, "#######0") & Format(MovCupomFiscal.Ordem, "00"), MovCupomFiscal.Ordem, MovCupomFiscal.CodigoProduto, MovCupomFiscal.Periodo) Then
                    Else
                        Call GravaAuditoria(1, Me.name, 25, "Não excluiu nota de abastecimento:" & MovCupomFiscal.NumeroCupom)
                        Call CriaLogCupom(Time & "ExcluiNotaAbastecimento: Não excluiu nota de abastecimento:" & MovCupomFiscal.NumeroCupom & " Data:" & MovCupomFiscal.Data & " Per:" & MovCupomFiscal.Periodo & " Produto:" & MovCupomFiscal.CodigoProduto)
                        MsgBox "Não foi possível excluir nota de abastecimento.", vbCritical, "Erro de Integridade!"
                    End If
                Else
                    Call GravaAuditoria(1, Me.name, 25, "Não localizou nota de abastecimento:" & MovCupomFiscal.NumeroCupom)
                    Call CriaLogCupom(Time & "ExcluiNotaAbastecimento: Não localizou nota de abastecimento:" & MovCupomFiscal.NumeroCupom & " Data:" & MovCupomFiscal.Data & " Per:" & MovCupomFiscal.Periodo & " Produto:" & MovCupomFiscal.CodigoProduto)
                    MsgBox "Não foi possível localizar nota de abastecimento.", vbCritical, "Erro de Integridade!"
                End If
            End If
        Else
            If i = 1 Then
                Call GravaAuditoria(1, Me.name, 25, "Não foi possível estornar nota de abastecimento no caixa.")
                MsgBox "Não foi possível estornar nota de abastecimento no caixa!", vbCritical, "Erro de Integridade!"
            End If
        End If
    Next
    Exit Sub

trata_erro:
    Call CriaLogCupom("Erro ExcluiNotaAbastecimento: Erro=" & Err.Number & " - " & Err.Description)
    Call GravaAuditoria(1, Me.name, 25, "ExcluiNotaAbastecimento: Erro inesperado...")
End Sub
Private Sub ExcluiSaidaBomba(x_produto As Long, x_quantidade As Currency)
    Dim xTipoCombustivel As String
    xTipoCombustivel = Bomba.LocalizarCodigoProduto(g_empresa, x_produto)
    If Combustivel.LocalizarCodigo(g_empresa, xTipoCombustivel) Then
        Combustivel.QuantidadeEmEstoque = Combustivel.QuantidadeEmEstoque + x_quantidade
        If Not Combustivel.Alterar(g_empresa, xTipoCombustivel) Then
            MsgBox "Não foi possível alterar registro de combustível!", vbInformation, "Erro de Integridade!"
        End If
    End If
End Sub
Private Sub ExcluiSaidaProduto(pCodigoProduto As Long, pQuantidade As Currency)
    If Estoque.LocalizarCodigo(g_empresa, pCodigoProduto) Then
        Estoque.Quantidade = Estoque.Quantidade + pQuantidade
        If Not Estoque.Alterar(g_empresa, pCodigoProduto) Then
            MsgBox "Não foi possível alterar o estoque!", vbInformation, "Erro de Integridade!"
        End If
    Else
        MsgBox "Estoque não cadastrado.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Function ExisteCupom() As Boolean
    Dim i As Integer

    On Error GoTo trata_erro
    
    ExisteCupom = False
    If MovCupomFiscal.LocalizarNumeroData(g_empresa, lNumeroPDV, CLng(lNumeroCupom), lData) Then
        ExisteCupom = True
        lPeriodo = MovCupomFiscal.Periodo
        cboTipoSubEstoque.ListIndex = -1
        For i = 0 To cboTipoSubEstoque.ListCount - 1
            If cboTipoSubEstoque.ItemData(i) = MovCupomFiscal.TipoSubEstoque Then
                cboTipoSubEstoque.ListIndex = i
                Exit For
            End If
        Next
        txt_cliente = MovCupomFiscal.CodigoCliente
        dtcboCliente.BoundText = MovCupomFiscal.CodigoCliente
    Else
        ExisteCupom = False
    End If
    Exit Function

trata_erro:
    Call CriaLogCupom("Erro ExisteCupom: Erro=" & Err.Number & " - " & Err.Description)
End Function
'ALEX - ABORTAR NFCE
Function RegistroNFCeAtualGravado() As CadastroDLL.cMovDocEletronicoCabecalho
    Dim i As Integer
    Dim xMovDocEletronicoCabecalho As New CadastroDLL.cMovDocEletronicoCabecalho

    On Error GoTo trata_erro
    
    RegistroNFCeAtualGravado = Nothing
    If xMovDocEletronicoCabecalho.LocalizarCodigo(g_empresa, lData, False, True, MODELO_NFCE, lSerieNFCe, CStr(lNumeroCupom)) Then
        RegistroNFCeAtualGravado = xMovDocEletronicoCabecalho
    End If
    
    Exit Function

trata_erro:
    Call CriaLogCupom("Erro RegistroNFCeAtualGravado: Erro=" & Err.Number & " - " & Err.Description)
End Function
'
'OBTEM DADOS DO CUPOM GRAVADO PARA GERAÇÃO DA NFCE
'
Private Sub EnviaDadosParaNFCe(pNumeroCupom As Long, pDataCupom As Date)
    Dim rsDadosParaNFCe As New ADODB.Recordset
    Dim xTipoServico As String
    Dim xTextoSolicitacao As String
    Dim xRealizarImpressao As String
    
    
On Error GoTo trata_erro

    If ConfiguracaoDiversa.LocalizarCodigo(g_empresa, "NFCe Imprimir Atraves") Then
        If UCase(ConfiguracaoDiversa.Texto) = "TECNOSPEED" Then
            xRealizarImpressao = "true"
        Else
            xRealizarImpressao = "false"
        End If
    End If


    xRealizarImpressao = IIf(lPOSImprimeNFCe = False, "false", xRealizarImpressao)

 '   MsgBox "EnviaDadosParaNFCe - ALEX TESTE 2"
    xTipoServico = "NFCe 3.10"
    
    If ConfiguracaoDiversa.LocalizarCodigo(1, "VERSAO NFCE") Then
        xTipoServico = "NFCe" & " " & ConfiguracaoDiversa.Texto
    End If
    
    Set rsDadosParaNFCe = ObtenhaDadosParaNFCEDocumentoEletronico(pNumeroCupom, pDataCupom)
  '  MsgBox "EnviaDadosParaNFCe - ALEX TESTE 3"
    
    If rsDadosParaNFCe.RecordCount > 0 Then
        rsDadosParaNFCe.MoveFirst
        
   '     MsgBox "EnviaDadosParaNFCe - ALEX TESTE 4"
        xTextoSolicitacao = MontaTextoCabecalhoSolicitacaoNFCE(rsDadosParaNFCe, xTipoServico)
        
        xTextoSolicitacao = xTextoSolicitacao & MontaTextoItensSolicitacaoNFCE(rsDadosParaNFCe)
        
    '    MsgBox "EnviaDadosParaNFCe - ALEX TESTE 5"
        If (AtualizaTabelaSolicitacaoNFCe(xTipoServico, "", xTextoSolicitacao, lNumeroCupom, "")) Then
        
          
           Dim Mensagem As String
     '       MsgBox "EnviaDadosParaNFCe - ALEX TESTE 6"
'           Set lProcessadorNFCE = New ProcessaNFCEFronteira
            'MsgBox "Vai chamar o Processamento - ALEX TESTE 7 - NSU = " & SolicitacaoFuncaoNFe.NSU
'           Mensagem = lProcessadorNFCE.ProcessaSolicitacaoFuncaoNFCe(SolicitacaoFuncaoNFe.NSU, SolicitacaoFuncaoNFe.CodigoEstabelecimento, GERADOR_NFCE_OOBJ)
            'MsgBox "EnviaDadosParaNFCe - ALEX TESTE 8 - Processamento Finalizado"
            
            If GravaDocumentoEletronicoEvento(MovDocEletronicoCabecalho, EVENTO_NFCE.FECHADA) Then
                MovDocEletronicoCabecalho.CodigoUltimoEvento = Val(EVENTO_NFCE.FECHADA)
                MovDocEletronicoCabecalho.ObservacaoEvento = "FECHADA"
            End If
            'gStringChamada = MovSolicitacaoFuncaoNFe.NSU_MovSolicitacaoFuncaoNFe & "|@|" & MovSolicitacaoFuncaoNFe.CodigoEstabelecimento_MovSolicitacaoFuncaoNFe & "|@|" & "1" & "|@|" '1-Oobj TXT, 2-Oobj XML, 3-cerrado
            'gStringChamada = MovSolicitacaoFuncaoNFe.NSU_MovSolicitacaoFuncaoNFe & "|@|" & MovSolicitacaoFuncaoNFe.CodigoEstabelecimento_MovSolicitacaoFuncaoNFe & "|@|" & "2" & "|@|" & xRealizarImpressao & "|@|" & gCNPJEmpresa & "|@|" & "True" & "|@|" & lNumeroCupom & "|@|" & MovDocEletronicoCabecalho.Serie & "|@|" & MovDocEletronicoCabecalho.DataEmissao & "|@|" & MovDocEletronicoCabecalho.Modelo & "|@|"   '1-Oobj TXT, 2-Oobj XML, 3-cerrado
            
            '4.0
            gStringChamada = MovSolicitacaoFuncaoNFe.NSU_MovSolicitacaoFuncaoNFe & "|@|" 'NSU
            gStringChamada = gStringChamada & MovSolicitacaoFuncaoNFe.CodigoEstabelecimento_MovSolicitacaoFuncaoNFe & "|@|" 'CODIGO DA EMPRESA
            gStringChamada = gStringChamada & "2" & "|@|" '1-Oobj TXT, 2-Oobj XML, 3-cerrado, 4-TECNOSPEED
            gStringChamada = gStringChamada & xRealizarImpressao & "|@|" 'se é pra realizar impressão
            gStringChamada = gStringChamada & gCNPJEmpresa & "|@|" 'cnpj do emitente
            gStringChamada = gStringChamada & "True" & "|@|" 'ProcessarRetorno
            gStringChamada = gStringChamada & lNumeroCupom & "|@|" 'Numero da NFCE
            gStringChamada = gStringChamada & MovDocEletronicoCabecalho.Serie & "|@|" 'Serie
            gStringChamada = gStringChamada & MovDocEletronicoCabecalho.DataEmissao & "|@|" 'Data Emissão
            gStringChamada = gStringChamada & MovDocEletronicoCabecalho.Modelo & "|@|" 'Modelo
            
            
            Call CriaLogCupom("[NFCE] Erro EnviaDadosParaNFCe: gStringChamada=" & gStringChamada)
            
            If ConfiguracaoDiversa.LocalizarCodigo(1, "PETROMOVELAUTO AUTORIZA NFCE") Then
                If ConfiguracaoDiversa.Verdadeiro = True Then
                    Call CriaLogCupom("[NFCE] Erro EnviaDadosParaNFCe: PETROMOVELAUTO AUTORIZA NFCE - ConfiguracaoDiversa.Verdadeiro=True")
                    'If VerificarAtivarPetromovelAuto Then
                        If Not GravaSolicitacaoProcessamentoNFCe(MovSolicitacaoFuncaoNFe.NumeroNFe_MovSolicitacaoFuncaoNFe, gStringChamada) Then
                            MsgBox "Não foi possível gravar a Solicitação do processamento para NFC-e!", vbCritical, "Erro de Integridade"
                        Else
                             Call AtivaDesativaAguarde("Aguarde! Processando NFCe... NSU(" & MovSolicitacaoFuncaoNFe.NSU_MovSolicitacaoFuncaoNFe & ")", True)
                             Call IniciaContadorAguarde(120)
                             lbl_mensagem.Caption = "Aguarde... Processando NFCe."
                             DoEvents
                        End If
                    'Else
                        'Call MovSolicitacaoFuncaoNFe.DefineHoraCancelamentoHost(MovSolicitacaoFuncaoNFe.NSU_MovSolicitacaoFuncaoNFe, Now, gVersaoSGP)
                    'End If

                    gStringChamada = ""
                Else
                    Call CriaLogCupom("[NFCE] Erro EnviaDadosParaNFCe: PETROMOVELAUTO AUTORIZA NFCE - ConfiguracaoDiversa.Verdadeiro=False")
                    Call menu_personalizado.GravaSgpNetCadastroIni("ProcessaNFCe")
                End If
            Else
                Call menu_personalizado.GravaSgpNetCadastroIni("ProcessaNFCe")
            End If
        End If
        
    End If
    Set rsDadosParaNFCe = Nothing
    Exit Sub

trata_erro:
    Dim ErroNFCE As String
    ErroNFCE = Err.Description
    Call CriaLogCupom("[NFCE] Erro EnviaDadosParaNFCe: Erro=" & Err.Number & " - " & ErroNFCE)
    MsgBox "Não foi possível gerar NFC-e. " & vbCrLf & ErroNFCE, vbCritical, "Erro Grave!"
    Exit Sub

End Sub

Private Function VerificarAtivarPetromovelAuto(ByVal pComunicarComBanco As Boolean) As Boolean
    VerificarAtivarPetromovelAuto = False
    
    If ComunicaPetromovelAutoBD = False Then
        Call CriaLogCupom("[VerificarAtivarPetromovelAuto] PetromovelAuto está Inativo.")
        'If (MsgBox("Não foi possível comunicar com o progama PetromovelAuto." & vbCrLf & vbCrLf & "Se este programa não for aberto, " & vbCrLf & "as funcionalidades da NFCe não serão executadas." & vbCrLf & vbCrLf & "Deseja abrir programa de emissão de NFCe?", vbQuestion + vbYesNo + vbDefaultButton1, "Deseja Abrir Programa PetromovelAuto?") = vbYes) Then
            Call CriaLogCupom("[VerificarAtivarPetromovelAuto] O SGP irá ativar o PetromovelAuto automaticamente")
            
            'VerificarAtivarPetromovelAuto = LoopAbrePetromovelAuto
            If LoopAbrePetromovelAuto(pComunicarComBanco) Then
                VerificarAtivarPetromovelAuto = True
            Else
                MsgBox "Não foi possível ativar PetromovelAuto!" & vbCrLf & "Contacte o suporte técnico.", vbCritical, "PetromovelAuto está inativo"
                VerificarAtivarPetromovelAuto = False
            End If
        'Else
          'Call CriaLogCupom("[VerificarAtivarPetromovelAuto] Operador optou por NÃO ativar o PetromovelAuto automaticamente")
        'End If
    Else
        VerificarAtivarPetromovelAuto = True
    End If

End Function

''Percorre impressoras instaladas no spooler do windows e define a primeira térmica encontrada com padrão
'Private Function DefineImpressoraTermicaComoPadrao() As Boolean
'
'DefineImpressoraTermicaComoPadrao = False
'Dim Impressora As Printer
'Dim Contador As Byte
'Contador = 0
'For Each Impressora In Printers
'    If UCase(Impressora.DeviceName) Like "*TM-T20*" Or UCase(Impressora.DeviceName) Like "*TM-T8*" Or UCase(Impressora.DeviceName) Like "*MP-4*" Or UCase(Impressora.DeviceName) Like "*MP-2*" Then
'        gIDImpressoraPadrao = Contador
'        Set Printer = Printers(gIDImpressoraPadrao)
'        DefineImpressoraTermicaComoPadrao = True
'        Exit Function
'    End If
'    Contador = Contador + 1
'Next
'MsgBox "Nenhuma Impressora térmica foi encontrada!"
'End Function
Private Sub ConsisteNFCePendentes(ByVal pMinutosDoTempoAtual As Integer)
    'pMinutosDoTempoAtual
    'esta variável diz para buscar NFCe no banco com x munutos a menos da hora atual
    gStringChamada = pMinutosDoTempoAtual & "|@|"
    Call menu_personalizado.GravaSgpNetCadastroIni("ConsisteNFCePendentes")
End Sub
''Percorre impressoras instaladas no spooler do windows e define a primeira térmica encontrada com padrão
'Private Function DefineImpressoraTermicaComoPadrao() As Boolean
'
'DefineImpressoraTermicaComoPadrao = False
'Dim Impressora As Printer
'Dim Contador As Byte
'Contador = 0
'For Each Impressora In Printers
'    If UCase(Impressora.DeviceName) Like "*TM-T20*" Or UCase(Impressora.DeviceName) Like "*TM-T8*" Or UCase(Impressora.DeviceName) Like "*MP-4*" Or UCase(Impressora.DeviceName) Like "*MP-2*" Then
'        gIDImpressoraPadrao = Contador
'        Set Printer = Printers(gIDImpressoraPadrao)
'        DefineImpressoraTermicaComoPadrao = True
'        Exit Function
'    End If
'    Contador = Contador + 1
'Next
'MsgBox "Nenhuma Impressora térmica foi encontrada!"
'End Function

Private Function GeraSolicitacaoImpressaoNFCE(ByVal pChaveNFCE As String, ByVal pNumeroDaNota As Long, ByVal pNumeroLote As String) As Boolean

    Dim rsDadosSolicitacao As New ADODB.Recordset
    
    GeraSolicitacaoImpressaoNFCE = False
    
    GeraSolicitacaoImpressaoNFCE = AtualizaTabelaSolicitacaoNFCe("IMPRESSAO", pChaveNFCE, Empty, pNumeroDaNota, pNumeroLote)
        

End Function
'HOJE-ALEX
Private Function ObtenhaDadosSolicitacaoNFCeEmitida(ByVal pNSU As Long, ByVal pIdEstabelecimento As Integer) As ADODB.Recordset
    On Error GoTo trata_erro
    
    Dim rsDadosSolicitacao As New ADODB.Recordset
    
    Dim xSQL As String
    
    
    xSQL = "SELECT NSU_MovSolicitacaoFuncaoNFe, ChaveAcessoNFe_MovSolicitacaoFuncaoNFe, CodigoRetorno_MovSolicitacaoFuncaoNFe, Mensagem_MovSolicitacaoFuncaoNFe, NumeroLote_MovSolicitacaoFuncaoNFe, NumeroNFe_MovSolicitacaoFuncaoNFe"
    xSQL = xSQL & " FROM MovSolicitacaoFuncaoNFe "
    xSQL = xSQL & " WHERE NSU_MovSolicitacaoFuncaoNFe = " & pNSU
    xSQL = xSQL & " AND CodigoEstabelecimento_MovSolicitacaoFuncaoNFe = " & pIdEstabelecimento
    
    'Abre RecordSet
    Set rsDadosSolicitacao = New ADODB.Recordset
    Set rsDadosSolicitacao = Conectar.RsConexao(xSQL)

Set ObtenhaDadosSolicitacaoNFCeEmitida = rsDadosSolicitacao

Exit Function

trata_erro:
    Dim ErroNFCE As String
    ErroNFCE = Err.Description
    Call CriaLogCupom("[NFCE] Erro ObtenhaDadosSolicitacaoNFCeParaImpressao: Erro=" & Err.Number & " - " & ErroNFCE)
    MsgBox "Não foi possível obter dados solicitação NFC-e para impressão. " & vbCrLf & ErroNFCE, vbCritical, "Erro Grave ao gerar impressão!"
    Exit Function


End Function
'INUTILIZAR
Private Function ObtenhaDadosParaNFCE_OLD(ByVal pNumeroCupom As Long, ByVal pDataCupom As Date) As ADODB.Recordset

    Dim rsDadosParaNFCe As New ADODB.Recordset
    
    Dim i As Integer
    Dim xSQL As String
    Dim xTextoSolicitacao As String
   
    'O campo Telefone está sendo preenchido com Encerrante Final para utilização na NFCE
    'O campo Numero do Cheque está sendo preenchido com valores concatenados da bomba, Bico e Tanque
    
    xSQL = ""
    xSQL = xSQL & "SELECT Movimento_Cupom_Fiscal.Empresa, Movimento_Cupom_Fiscal.[Numero do Cupom],"
    xSQL = xSQL & "       Movimento_Cupom_Fiscal.[Data do Cupom], Movimento_Cupom_Fiscal.Ordem,"
    xSQL = xSQL & "       Movimento_Cupom_Fiscal.Hora, Movimento_Cupom_Fiscal.Periodo,"
    xSQL = xSQL & "       Movimento_Cupom_Fiscal.[Tipo do Movimento], Movimento_Cupom_Fiscal.[Codigo do Cliente],"
    xSQL = xSQL & "       Movimento_Cupom_Fiscal.[Codigo do Conveniado], Movimento_Cupom_Fiscal.[Codigo do Produto],"
    xSQL = xSQL & "       Movimento_Cupom_Fiscal.[Valor Unitario], Movimento_Cupom_Fiscal.Quantidade,"
    xSQL = xSQL & "       Movimento_Cupom_Fiscal.[Valor Total], Movimento_Cupom_Fiscal.[Forma de Pagamento],"
    xSQL = xSQL & "       Movimento_Cupom_Fiscal.[Valor Recebido], Movimento_Cupom_Fiscal.[Valor do Desconto],"
    xSQL = xSQL & "       Movimento_Cupom_Fiscal.[Valor Desconto Embutido], Movimento_Cupom_Fiscal.[Numero do Cheque],"
    xSQL = xSQL & "       Movimento_Cupom_Fiscal.[Operador], Movimento_Cupom_Fiscal.[Codigo da Aliquota],"
    xSQL = xSQL & "       Movimento_Cupom_Fiscal.Nome, Movimento_Cupom_Fiscal.Telefone, "
    xSQL = xSQL & "       Movimento_Cupom_Fiscal.[CPF CNPJ], Movimento_Cupom_Fiscal.[Tipo de Combustivel],"
    xSQL = xSQL & "       Movimento_Cupom_Fiscal.[Codigo da Ecf], Movimento_Cupom_Fiscal.[Codigo do Grupo],"
    xSQL = xSQL & "       Movimento_Cupom_Fiscal.[Tipo do SubEstoque],"

    xSQL = xSQL & "       Movimento_Cupom_Fiscal_Item.[Valor do Desconto], Movimento_Cupom_Fiscal_Item.[Valor do Acrescimo],"

    xSQL = xSQL & "       Produto.Nome AS NomeProduto, Produto.Unidade, Produto.[Codigo de Barra],"
    xSQL = xSQL & "       Produto.[CST ICMS], Produto.[CST PIS], Produto.[CST COFINS],"
    xSQL = xSQL & "       Produto.[Codigo NCM], Produto.[Codigo ANP], Produto.CEST,"
    xSQL = xSQL & "       Aliquota.[Codigo Fiscal] , Aliquota.[Aliquota do Imposto]"
    xSQL = xSQL & "  FROM [movimento_cupom_fiscal], Movimento_Cupom_Fiscal_Item, Produto, Aliquota"
    xSQL = xSQL & " WHERE movimento_cupom_fiscal.Empresa = " & g_empresa
    xSQL = xSQL & "   AND Movimento_Cupom_Fiscal.[Data do Cupom] = " & preparaData(pDataCupom)
    xSQL = xSQL & "   AND Movimento_Cupom_Fiscal.[Numero do Cupom] = " & pNumeroCupom

    xSQL = xSQL & "   AND Movimento_Cupom_Fiscal.[Cupom Cancelado] = " & preparaBooleano(False)
    xSQL = xSQL & "   AND Movimento_Cupom_Fiscal.[Item Cancelado] = " & preparaBooleano(False)

    xSQL = xSQL & "   AND Movimento_Cupom_Fiscal.Empresa = Movimento_Cupom_Fiscal_Item.Empresa"
    xSQL = xSQL & "   AND Movimento_Cupom_Fiscal.[Data do Cupom] = Movimento_Cupom_Fiscal_Item.Data"
    xSQL = xSQL & "   AND Movimento_Cupom_Fiscal.[Numero do Cupom] = Movimento_Cupom_Fiscal_Item.[Numero do Cupom]"
    xSQL = xSQL & "   AND Movimento_Cupom_Fiscal.Ordem = Movimento_Cupom_Fiscal_Item.Ordem"

    xSQL = xSQL & "   AND Movimento_Cupom_Fiscal.[Codigo do Produto] = Produto.Codigo"

    xSQL = xSQL & "   AND Movimento_Cupom_Fiscal.[Codigo da Aliquota] = Aliquota.Codigo"
    xSQL = xSQL & "   AND Aliquota.[Serie ECF] = " & preparaTexto(lSerieECF)

    xSQL = xSQL & " ORDER BY Movimento_Cupom_Fiscal.Ordem"

    'Abre RecordSet
    Set rsDadosParaNFCe = New ADODB.Recordset
    Set rsDadosParaNFCe = Conectar.RsConexao(xSQL)
    
    
    Set ObtenhaDadosParaNFCE_OLD = rsDadosParaNFCe

End Function
Private Sub CancelarFinalizacaoNFCe()
    'frm_fechamento_cupom.Width = 5700
    'frm_fechamento_cupom.Top = 500
    'frm_fechamento_cupom.Left = 120
    frm_fechamento_cupom.Top = 200
    frm_fechamento_cupom.Left = 200
    frm_fechamento_cupom.Height = 1
    frm_fechamento_cupom.Width = 1
    frm_fechamento_cupom.ZOrder 1
    frm_fechamento_cupom.Enabled = False
    cmdIniciaProcessoFinalizacaoNFCe.Visible = True
    frmDados.Enabled = True
    NovoCupom
End Sub
Private Function ObtenhaDadosParaNFCEDocumentoEletronico(ByVal pNumeroNFCe As Long, ByVal pDataEmissao As Date) As ADODB.Recordset

    Dim rsDadosParaNFCe As New ADODB.Recordset
    
    Dim i As Integer
    Dim xSQL As String
    Dim xTextoSolicitacao As String
   
    'O campo Telefone está sendo preenchido com Encerrante Final para utilização na NFCE
    'O campo Numero do Cheque está sendo preenchido com valores concatenados da bomba, Bico e Tanque
    
    
    
    xSQL = ""
    xSQL = xSQL & "SELECT IdEstabelecimento_MovDECabecalho, Numero_MovDECabecalho,"
    xSQL = xSQL & "DataEmissao_MovDECabecalho,"
    xSQL = xSQL & "Ordem_MovDEItem,"
    xSQL = xSQL & "HoraSaida_MovDECabecalho,"
    xSQL = xSQL & "IdClienteFornecedor_MovDECabecalho,"
    xSQL = xSQL & "IdProduto_MovDEItem,"
    xSQL = xSQL & "ValorUnitario_MovDEItem,"
    xSQL = xSQL & "Quantidade_MovDEItem,"
    xSQL = xSQL & "ROUND(ValorUnitario_MovDEItem * Quantidade_MovDEItem, 4) AS TotalBruto_MovDEItem,"
    xSQL = xSQL & "ValorTotal_MovDECabecalho,"
    xSQL = xSQL & "FormaPagamento_MovDECabecalho,"
    xSQL = xSQL & "TipoSubEstoque_MovDECabecalho,"
    xSQL = xSQL & "ValorTotalLiquido_MovDEItem,"
    xSQL = xSQL & "ValorDesconto_MovDECabecalho,"
    xSQL = xSQL & "IdUsuario_MovDECabecalho,"
    xSQL = xSQL & "ValorDesconto_MovDEItem,"
    xSQL = xSQL & "EncerranteFinal_MovDEItem,"
    xSQL = xSQL & "NumeroBomba_MovDEItem,"
    xSQL = xSQL & "NumeroBico_MovDEItem,"
    xSQL = xSQL & "NumeroTanque_MovDEItem,"
    xSQL = xSQL & "TipoCombustivel_MovDEItem,"
    xSQL = xSQL & "CFOP_MovDEItem,"
    xSQL = xSQL & "Produto.Nome As NomeProduto,"
    xSQL = xSQL & "Produto.Unidade,"
    xSQL = xSQL & "Produto.[Codigo de Barra],"
    xSQL = xSQL & "ValorBcIcms_MovDEItem,"
    xSQL = xSQL & "ValorIcms_MovDEItem,"
    xSQL = xSQL & "CstIcms_MovDEItem," '"Produto.[CST ICMS],"
    xSQL = xSQL & "ValorBcIcmsSt_MovDEItem,"
    xSQL = xSQL & "ValorIcmsSt_MovDEItem,"
    xSQL = xSQL & "AliquotaIcmsSt_MovDEItem,"
    xSQL = xSQL & "CstPis_MovDEItem," '"Produto.[CST PIS],"
    xSQL = xSQL & "ValorBcPis_MovDEItem,"
    xSQL = xSQL & "ValorBcCofins_MovDEItem,"
    xSQL = xSQL & "CstCofins_MovDEItem," '"Produto.[CST COFINS],"
    xSQL = xSQL & "AliquotaPis_MovDEItem,"
    xSQL = xSQL & "AliquotaCofins_MovDEItem,"
    xSQL = xSQL & "ValorPis_MovDEItem,"
    xSQL = xSQL & "ValorCofins_MovDEItem,"
    xSQL = xSQL & "Produto.[Codigo NCM],"
    xSQL = xSQL & "Produto.[Codigo ANP],"
    xSQL = xSQL & "Produto.CEST,"
    xSQL = xSQL & "Aliquota.[Codigo Fiscal] ,"
    xSQL = xSQL & "AliquotaIcms_MovDEItem" '"Aliquota.[Aliquota do Imposto]"
    xSQL = xSQL & " FROM Produto, MovimentoDocumentoEletronicoCabecalho,"
    xSQL = xSQL & " Aliquota, MovimentoDocumentoEletronicoItem"
    xSQL = xSQL & " WHERE IdEstabelecimento_MovDECabecalho =" & g_empresa
    xSQL = xSQL & " AND DataEmissao_MovDECabecalho = " & preparaData(pDataEmissao)
    xSQL = xSQL & " AND Numero_MovDECabecalho = " & pNumeroNFCe
    xSQL = xSQL & " AND Cancelado_MovDECabecalho = " & preparaBooleano(False)
    xSQL = xSQL & " AND Cancelado_MovDEItem = " & preparaBooleano(False)
    xSQL = xSQL & " AND IdEstabelecimento_MovDECabecalho = IdEstabelecimento_MovDEItem   "
    xSQL = xSQL & " AND DataEmissao_MovDECabecalho = DataEmissao_MovDEItem   "
    xSQL = xSQL & " AND Numero_MovDECabecalho = Numero_MovDEItem   "
    xSQL = xSQL & " AND IdProduto_MovDEItem = Produto.Codigo   "
    xSQL = xSQL & " AND Produto.[Codigo da Aliquota] = Aliquota.Codigo   "
    xSQL = xSQL & " AND Aliquota.[Serie ECF] = " & preparaTexto(lSerieECF)
    xSQL = xSQL & " AND Entrada_MovDECabecalho = " & preparaBooleano(False)
    xSQL = xSQL & " AND Saida_MovDECabecalho = " & preparaBooleano(True)
    xSQL = xSQL & " AND Modelo_MovDECabecalho = " & preparaTexto(MODELO_NFCE)
    xSQL = xSQL & " AND Serie_MovDECabecalho = " & preparaTexto(lSerieNFCe)
    xSQL = xSQL & "ORDER BY Ordem_MovDEItem"
    

    'Abre RecordSet
    Set rsDadosParaNFCe = New ADODB.Recordset
    Set rsDadosParaNFCe = Conectar.RsConexao(xSQL)
    CriaLogCupom ("[ObtenhaDadosParaNFCEDocumentoEletronico] - " & xSQL)
    
    Set ObtenhaDadosParaNFCEDocumentoEletronico = rsDadosParaNFCe

End Function
Private Sub VerificaDescontoPersonalizado()
    
    On Error GoTo trata_erro
    
    lTotalItem(lOrdem + 1) = fValidaValor(txt_valor_total.Text)
    lValorUnitarioSemDesconto = fValidaValor(txt_valor_unitario.Text)
    lValorTotalSemDesconto = fValidaValor(txt_valor_total.Text)
    If txt_produto.Text <> "" Then
        If Estoque.LocalizarCodigo(g_empresa, CLng(txt_produto.Text)) Then
            lValorUnitarioSemDesconto = Estoque.PrecoVenda
            lValorTotalSemDesconto = Format(Estoque.PrecoVenda * fValidaValor(txt_quantidade.Text), "###,##0.00")
            lTotalItem(lOrdem + 1) = lValorTotalSemDesconto
        End If
    End If
    'Verifica desconto personalizado
    lDescontoItemEmbutido = 0
    lAcrescimoItemEmbutido = 0
    Call CriaLogMonitoraBug("Monitora Desconto Personalizado", "Movimento_NFce_Auto", "VerificaDescontoPersonalizado -> As variáveis lDescontoItemEmbutido e lAcrescimoItemEmbutido foram zeradas")
    If Val(l_codigo_cliente) > 0 Then
        Call CriaLogMonitoraBug("Monitora Desconto Personalizado", "Movimento_NFce_Auto", "Código do Cliente: " & l_codigo_cliente)
        If MovDescontoPersonalizado.LocalizarCodigo(l_codigo_cliente, CLng(dtcboProduto.BoundText)) Then
            Call CriaLogMonitoraBug("Monitora Desconto Personalizado", "Movimento_NFce_Auto", "Cliente tem preço personalizado. Produto: " & dtcboProduto.BoundText)
            Call GravaAuditoria(1, Me.name, 22, "Preço diferenciado Cli:" & txt_cliente.Text & " Prod:" & txt_produto.Text)
            MsgBox "Este cliente tem preço diferenciado." & Chr(10) & Chr(10) & "O sistema irá calcular automaticamente!", vbInformation, "Preço Diferenciado para o Cliente!"
            If MovDescontoPersonalizado.PrecoFixo > 0 Then
                Call CriaLogMonitoraBug("Monitora Desconto Personalizado", "Movimento_NFce_Auto", "PrecoFixo: " & MovDescontoPersonalizado.PrecoFixo)
                'Valor Fixo
                If MovDescontoPersonalizado.PrecoFixo < fValidaValor(txt_valor_unitario.Text) Then
                    'Desconto
                    'lDescontoItemEmbutido = MovDescontoPersonalizado.PrecoFixo - fValidaValor(txt_valor_unitario.Text)
                    lDescontoItemEmbutido = MovDescontoPersonalizado.PrecoFixo - lValorUnitarioSemDesconto
                Else
                    'Acréscimo
                    'lAcrescimoItemEmbutido = fValidaValor(txt_valor_unitario.Text) - MovDescontoPersonalizado.PrecoFixo
                    lAcrescimoItemEmbutido = lValorUnitarioSemDesconto - MovDescontoPersonalizado.PrecoFixo
                End If
                'Define Valor Fixo
                txt_valor_unitario.Text = Format(MovDescontoPersonalizado.PrecoFixo, "###,###,##0.0000")
                txt_valor_total.Text = Format(fValidaValor4(txt_valor_unitario.Text) * fValidaValor(txt_quantidade.Text), "###,##0.00")
            ElseIf MovDescontoPersonalizado.Desconto = True Then
                Call CriaLogMonitoraBug("Monitora Desconto Personalizado", "Movimento_NFce_Auto", "Desconto = True")
                'Calcula Desconto
                If MovDescontoPersonalizado.ValoraDescontar > 0 Then
                    Call CriaLogMonitoraBug("Monitora Desconto Personalizado", "Movimento_NFce_Auto", "ValoraDescontar: " & MovDescontoPersonalizado.ValoraDescontar)
                    txt_valor_unitario.Text = Format(fValidaValor4(txt_valor_unitario.Text) - MovDescontoPersonalizado.ValoraDescontar, "###,###,##0.0000")
                    lDescontoItemEmbutido = fValidaValor(txt_valor_total.Text)
                    txt_valor_total.Text = Format(fValidaValor4(txt_valor_unitario.Text) * fValidaValor(txt_quantidade.Text), "###,##0.00")
                    lDescontoItemEmbutido = lDescontoItemEmbutido - fValidaValor(txt_valor_total.Text)
                Else
                    Call CriaLogMonitoraBug("Monitora Desconto Personalizado", "Movimento_NFce_Auto", "PercentualaDescontar: " & MovDescontoPersonalizado.PercentualaDescontar)
                    txt_valor_unitario.Text = Format(fValidaValor4(txt_valor_unitario.Text) - Format((fValidaValor4(txt_valor_unitario.Text) * MovDescontoPersonalizado.PercentualaDescontar / 100), "00000000.0000"), "###,###,##0.0000")
                    lDescontoItemEmbutido = fValidaValor(txt_valor_total.Text)
                    txt_valor_total.Text = Format(fValidaValor4(txt_valor_unitario.Text) * fValidaValor(txt_quantidade.Text), "###,##0.00")
                    lDescontoItemEmbutido = lDescontoItemEmbutido - fValidaValor(txt_valor_total.Text)
                End If
            Else
                Call CriaLogMonitoraBug("Monitora Desconto Personalizado", "Movimento_NFce_Auto", "Desconto = False")
                'Calcula Acréscimo
                If MovDescontoPersonalizado.ValoraDescontar > 0 Then
                    Call CriaLogMonitoraBug("Monitora Desconto Personalizado", "Movimento_NFce_Auto", "Valor a acrescentar: " & MovDescontoPersonalizado.ValoraDescontar)
                    txt_valor_unitario.Text = Format(fValidaValor4(txt_valor_unitario.Text) + MovDescontoPersonalizado.ValoraDescontar, "###,###,##0.0000")
                    lAcrescimoItemEmbutido = fValidaValor(txt_valor_total.Text)
                    txt_valor_total.Text = Format(fValidaValor4(txt_valor_unitario.Text) * fValidaValor(txt_quantidade.Text), "###,##0.00")
                    lAcrescimoItemEmbutido = fValidaValor(txt_valor_total.Text) - lAcrescimoItemEmbutido
                Else
                    Call CriaLogMonitoraBug("Monitora Desconto Personalizado", "Movimento_NFce_Auto", "Percentual a acrescentar: " & MovDescontoPersonalizado.PercentualaDescontar)
                    txt_valor_unitario.Text = Format(fValidaValor4(txt_valor_unitario.Text) + Format((fValidaValor4(txt_valor_unitario.Text) * MovDescontoPersonalizado.PercentualaDescontar / 100), "00000000.0000"), "###,###,##0.0000")
                    lAcrescimoItemEmbutido = fValidaValor(txt_valor_total.Text)
                    txt_valor_total.Text = Format(fValidaValor4(txt_valor_unitario.Text) * fValidaValor(txt_quantidade.Text), "###,##0.00")
                    lAcrescimoItemEmbutido = fValidaValor(txt_valor_total.Text) - lAcrescimoItemEmbutido
                End If
            End If
        Else
            'Desconto por Grupo de Cliente
            If MovDescontoGrupoCliente.LocalizarCodigo(Cliente.CodigoGrupoCliente, CLng(dtcboProduto.BoundText)) Then
                Call CriaLogMonitoraBug("Monitora Desconto Personalizado", "Movimento_NFce_Auto", "Cliente tem Gruo de preço personalizado. Produto: " & dtcboProduto.BoundText & " - Grupo: " & MovDescontoGrupoCliente.CodigoGrupoCliente)
                Call GravaAuditoria(1, Me.name, 22, "Preço diferenciado Por Grupo de Cliente:" & Cliente.CodigoGrupoCliente)
                Call GravaAuditoria(1, Me.name, 22, "Preço diferenciado Cli:" & txt_cliente.Text & " Prod:" & txt_produto.Text)
                MsgBox "Este cliente tem preço diferenciado." & Chr(10) & Chr(10) & "O sistema irá calcular automaticamente!", vbInformation, "Preço Diferenciado para o Cliente!"
                If MovDescontoGrupoCliente.PrecoFixo > 0 Then
                    Call CriaLogMonitoraBug("Monitora Desconto Personalizado", "Movimento_NFce_Auto", "Grupo PrecoFixo: " & MovDescontoGrupoCliente.PrecoFixo)
                    'Valor Fixo
                    If MovDescontoGrupoCliente.PrecoFixo < fValidaValor(txt_valor_unitario.Text) Then
                        'Desconto
                        'lDescontoItemEmbutido = MovDescontoGrupoCliente.PrecoFixo - fValidaValor(txt_valor_unitario.Text)
                        lDescontoItemEmbutido = MovDescontoGrupoCliente.PrecoFixo - lValorUnitarioSemDesconto
                    Else
                        'Acréscimo
                        'lAcrescimoItemEmbutido = fValidaValor(txt_valor_unitario.Text) - MovDescontoGrupoCliente.PrecoFixo
                        lAcrescimoItemEmbutido = lValorUnitarioSemDesconto - MovDescontoGrupoCliente.PrecoFixo
                    End If
                    'Define Valor Fixo
                    txt_valor_unitario.Text = Format(MovDescontoGrupoCliente.PrecoFixo, "###,###,##0.0000")
                    txt_valor_total.Text = Format(fValidaValor4(txt_valor_unitario.Text) * fValidaValor(txt_quantidade.Text), "###,##0.00")
                ElseIf MovDescontoGrupoCliente.Desconto = True Then
                    Call CriaLogMonitoraBug("Monitora Desconto Personalizado", "Movimento_NFce_Auto", "Grupo Desconto = True")
                    'Calcula Desconto
                    If MovDescontoGrupoCliente.ValoraDescontar > 0 Then
                        txt_valor_unitario.Text = Format(fValidaValor4(txt_valor_unitario.Text) - MovDescontoGrupoCliente.ValoraDescontar, "###,###,##0.0000")
                        lDescontoItemEmbutido = fValidaValor(txt_valor_total.Text)
                        txt_valor_total.Text = Format(fValidaValor4(txt_valor_unitario.Text) * fValidaValor(txt_quantidade.Text), "###,##0.00")
                        lDescontoItemEmbutido = lDescontoItemEmbutido - fValidaValor(txt_valor_total.Text)
                    Else
                        txt_valor_unitario.Text = Format(fValidaValor4(txt_valor_unitario.Text) - Format((fValidaValor4(txt_valor_unitario.Text) * MovDescontoGrupoCliente.PercentualaDescontar / 100), "00000000.0000"), "###,###,##0.0000")
                        lDescontoItemEmbutido = fValidaValor(txt_valor_total.Text)
                        txt_valor_total.Text = Format(fValidaValor4(txt_valor_unitario.Text) * fValidaValor(txt_quantidade.Text), "###,##0.00")
                        lDescontoItemEmbutido = lDescontoItemEmbutido - fValidaValor(txt_valor_total.Text)
                    End If
                Else
                    Call CriaLogMonitoraBug("Monitora Desconto Personalizado", "Movimento_NFce_Auto", "Grupo Desconto = False")
                    'Calcula Acréscimo
                    If MovDescontoGrupoCliente.ValoraDescontar > 0 Then
                        txt_valor_unitario.Text = Format(fValidaValor4(txt_valor_unitario.Text) + MovDescontoGrupoCliente.ValoraDescontar, "###,###,##0.0000")
                        lAcrescimoItemEmbutido = fValidaValor(txt_valor_total.Text)
                        txt_valor_total.Text = Format(fValidaValor4(txt_valor_unitario.Text) * fValidaValor(txt_quantidade.Text), "###,##0.00")
                        lAcrescimoItemEmbutido = fValidaValor(txt_valor_total.Text) - lAcrescimoItemEmbutido
                    Else
                        txt_valor_unitario.Text = Format(fValidaValor4(txt_valor_unitario.Text) + Format((fValidaValor4(txt_valor_unitario.Text) * MovDescontoGrupoCliente.PercentualaDescontar / 100), "00000000.0000"), "###,###,##0.0000")
                        lAcrescimoItemEmbutido = fValidaValor(txt_valor_total.Text)
                        txt_valor_total.Text = Format(fValidaValor4(txt_valor_unitario.Text) * fValidaValor(txt_quantidade.Text), "###,##0.00")
                        lAcrescimoItemEmbutido = fValidaValor(txt_valor_total.Text) - lAcrescimoItemEmbutido
                    End If
                End If
            End If
        End If
    End If
    Exit Sub

trata_erro:
    Call CriaLogMonitoraBug("Monitora Desconto Personalizado", "Movimento_NFce_Auto", "VerificaDescontoPersonalizado -> Erro=" & Err.Number & " - " & Err.Description)
    Call CriaLogCupom("Erro VerificaDescontoPersonalizado: Erro=" & Err.Number & " - " & Err.Description)
End Sub
Private Function DescontoPersonalizado(ByVal pCodigoCliente As Long, ByVal pCodigoGrupoCliente As Long, ByVal pCodigoProduto As Long, ByVal pValorUnitario As Currency) As Currency

    On Error GoTo trata_erro
    
    'Verifica desconto personalizado para gravar na nota
    DescontoPersonalizado = 0
    Call CriaLogMonitoraBug("Monitora Desconto Personalizado", "Movimento_NFce_Auto", "DescontoPersonalizado -> DescontoPersonalizado = 0")
    If Estoque.LocalizarCodigo(g_empresa, pCodigoProduto) Then
        pValorUnitario = Estoque.PrecoVenda
        Call CriaLogMonitoraBug("Monitora Desconto Personalizado", "Movimento_NFce_Auto", "pValorUnitario: " & pValorUnitario)
    Else
        Call CriaLogMonitoraBug("Monitora Desconto Personalizado", "Movimento_NFce_Auto", "Estoque.LocalizarCodigo = FALSE")
    End If
    If MovDescontoPersonalizado.LocalizarCodigo(pCodigoCliente, pCodigoProduto) Then
        Call CriaLogMonitoraBug("Monitora Desconto Personalizado", "Movimento_NFce_Auto", "Encontrado desconto personalizado pro pCodigoCliente: " & pCodigoCliente & " pCodigoProduto: " & pCodigoProduto)
        If MovDescontoPersonalizado.PrecoFixo > 0 Then
            'Valor Fixo
            Call CriaLogMonitoraBug("Monitora Desconto Personalizado", "Movimento_NFce_Auto", "MovDescontoPersonalizado.PrecoFixo: " & MovDescontoPersonalizado.PrecoFixo)
            If MovDescontoPersonalizado.PrecoFixo < pValorUnitario Then
                'Desconto
                DescontoPersonalizado = pValorUnitario - MovDescontoPersonalizado.PrecoFixo
            Else
                'Acréscimo
                DescontoPersonalizado = pValorUnitario - MovDescontoPersonalizado.PrecoFixo
            End If
            'Define Valor Fixo
        ElseIf MovDescontoPersonalizado.Desconto = True Then
            'Calcula Desconto
            Call CriaLogMonitoraBug("Monitora Desconto Personalizado", "Movimento_NFce_Auto", "MovDescontoPersonalizado.Desconto = TRUE")
            If MovDescontoPersonalizado.ValoraDescontar > 0 Then
                Call CriaLogMonitoraBug("Monitora Desconto Personalizado", "Movimento_NFce_Auto", "MovDescontoPersonalizado.ValoraDescontar: " & MovDescontoPersonalizado.ValoraDescontar)
                DescontoPersonalizado = MovDescontoPersonalizado.ValoraDescontar
            Else
                Call CriaLogMonitoraBug("Monitora Desconto Personalizado", "Movimento_NFce_Auto", "MovDescontoPersonalizado.PercentualaDescontar: " & MovDescontoPersonalizado.PercentualaDescontar)
                DescontoPersonalizado = Format(pValorUnitario * MovDescontoPersonalizado.PercentualaDescontar / 100, "00000000.0000")
            End If
        Else
            'Calcula Acréscimo
            Call CriaLogMonitoraBug("Monitora Desconto Personalizado", "Movimento_NFce_Auto", "MovDescontoPersonalizado.Desconto = FALSE")
            If MovDescontoPersonalizado.ValoraDescontar > 0 Then
                Call CriaLogMonitoraBug("Monitora Desconto Personalizado", "Movimento_NFce_Auto", "ACRESCENTAR MovDescontoPersonalizado.ValoraDescontar: " & MovDescontoPersonalizado.ValoraDescontar)
                DescontoPersonalizado = -MovDescontoPersonalizado.ValoraDescontar
            Else
                Call CriaLogMonitoraBug("Monitora Desconto Personalizado", "Movimento_NFce_Auto", "ACRESCENTAR MovDescontoPersonalizado.PercentualaDescontar: " & MovDescontoPersonalizado.PercentualaDescontar)
                DescontoPersonalizado = -(Format(pValorUnitario * MovDescontoPersonalizado.PercentualaDescontar / 100, "00000000.0000"))
            End If
        End If
    Else
    'Desconto Por Grupo de Cliente
        If MovDescontoGrupoCliente.LocalizarCodigo(pCodigoGrupoCliente, pCodigoProduto) Then
            Call CriaLogMonitoraBug("Monitora Desconto Personalizado", "Movimento_NFce_Auto", "Encontrado desconto MovDescontoGrupoCliente pro pCodigoGrupoCliente: " & pCodigoGrupoCliente & " pCodigoProduto: " & pCodigoProduto)
            If MovDescontoGrupoCliente.PrecoFixo > 0 Then
                'Valor Fixo
                Call CriaLogMonitoraBug("Monitora Desconto Personalizado", "Movimento_NFce_Auto", "MovDescontoGrupoCliente.PrecoFixo: " & MovDescontoGrupoCliente.PrecoFixo)
                If MovDescontoGrupoCliente.PrecoFixo < pValorUnitario Then
                    'Desconto
                    DescontoPersonalizado = pValorUnitario - MovDescontoGrupoCliente.PrecoFixo
                Else
                    'Acréscimo
                    DescontoPersonalizado = pValorUnitario - MovDescontoGrupoCliente.PrecoFixo
                End If
                'Define Valor Fixo
            ElseIf MovDescontoGrupoCliente.Desconto = True Then
                'Calcula Desconto
                Call CriaLogMonitoraBug("Monitora Desconto Personalizado", "Movimento_NFce_Auto", "MovDescontoGrupoCliente.Desconto = TRUE")
                If MovDescontoGrupoCliente.ValoraDescontar > 0 Then
                    Call CriaLogMonitoraBug("Monitora Desconto Personalizado", "Movimento_NFce_Auto", "MovDescontoGrupoCliente.ValoraDescontar: " & MovDescontoGrupoCliente.ValoraDescontar)
                    DescontoPersonalizado = MovDescontoGrupoCliente.ValoraDescontar
                Else
                    Call CriaLogMonitoraBug("Monitora Desconto Personalizado", "Movimento_NFce_Auto", "MovDescontoGrupoCliente.PercentualaDescontar: " & MovDescontoGrupoCliente.PercentualaDescontar)
                    DescontoPersonalizado = Format(pValorUnitario * MovDescontoGrupoCliente.PercentualaDescontar / 100, "00000000.0000")
                End If
            Else
                'Calcula Acréscimo
                Call CriaLogMonitoraBug("Monitora Desconto Personalizado", "Movimento_NFce_Auto", "MovDescontoGrupoCliente.Desconto = FALSE")
                If MovDescontoGrupoCliente.ValoraDescontar > 0 Then
                    Call CriaLogMonitoraBug("Monitora Desconto Personalizado", "Movimento_NFce_Auto", "ACRESCENTAR MovDescontoGrupoCliente.ValoraDescontar: " & MovDescontoGrupoCliente.ValoraDescontar)
                    DescontoPersonalizado = -MovDescontoGrupoCliente.ValoraDescontar
                Else
                    Call CriaLogMonitoraBug("Monitora Desconto Personalizado", "Movimento_NFce_Auto", "ACRESCENTAR MovDescontoGrupoCliente.PercentualaDescontar: " & MovDescontoGrupoCliente.PercentualaDescontar)
                    DescontoPersonalizado = -(Format(pValorUnitario * MovDescontoGrupoCliente.PercentualaDescontar / 100, "00000000.0000"))
                End If
            End If
        End If
    End If
    Exit Function
    
trata_erro:
    Call CriaLogMonitoraBug("Monitora Desconto Personalizado", "Movimento_NFce_Auto", "DescontoPersonalizado -> Erro=" & Err.Number & " - " & Err.Description)
    Call CriaLogCupom("Erro DescontoPersonalizado: Erro=" & Err.Number & " - " & Err.Description)
End Function
Private Sub VerificaSeExisteCupom()
    If MovCupomFiscal.LocalizarCodigo(g_empresa, lNumeroPDV, CLng(lNumeroCupom), lData, Val(lOrdem)) Then
        lNumeroCupom = MovCupomFiscal.NumeroCupom + 1 'NFCE - ALEX
        'MsgBox "Cupom Fiscal Existente!", vbInformation, "Erro de Integridade!" NFCE - ALEX
'        If MovCupomFiscal.Excluir(g_empresa, lNumeroPDV, CLng(lNumeroCupom), lData, Val(lOrdem)) Then
'            If Not MovCupomFiscalItem.Excluir(g_empresa, lNumeroPDV, lData, CLng(lNumeroCupom), Val(lOrdem)) Then
'                MsgBox "Não foi possível excluir o cupom fiscal!", vbInformation, "Erro de Integridade!"
'            End If
'        Else
'            MsgBox "Não foi possível excluir o cupom fiscal!", vbInformation, "Erro de Integridade!"
'        End If
    End If
End Sub
Private Sub ZZTotalizaCupomAbertoNoBanco()
    Dim xNumeroEcf As Integer
    Dim xNumeroCupom As Long
    Dim xData As Date
    Dim xOrdem As Integer
    Dim xValor As Currency
    Dim rstMovCupomFiscal As ADODB.Recordset
    Dim xSQL As String
    
    xSQL = "SELECT [Codigo da Ecf], Data, [Numero do Cupom], Quantidade, [Valor Total]"
    xSQL = xSQL & "  FROM Movimento_Cupom_Fiscal"
    xSQL = xSQL & " WHERE Empresa = " & g_empresa
    xSQL = xSQL & "   AND Data >= " & preparaData(CDate("01/07/2011"))
    xSQL = xSQL & "   AND [Forma de Pagamento] = 0"
    xSQL = xSQL & " ORDER BY [Codigo da Ecf], Data, [Numero do Cupom]"
    Set rstMovCupomFiscal = Conectar.RsConexao(xSQL)
    If rstMovCupomFiscal.RecordCount > 0 Then
        MsgBox "Será corrigido " & rstMovCupomFiscal.RecordCount & " cupons com forma de pagamento zerada!", vbInformation, "Atenção!"
        Do Until rstMovCupomFiscal.EOF
            xData = rstMovCupomFiscal!Data
            xNumeroEcf = rstMovCupomFiscal![Codigo da Ecf]
            xNumeroCupom = rstMovCupomFiscal![Numero do Cupom]
            xOrdem = 0
            xValor = 0
            Do Until MovCupomFiscal.LocalizarNumeroProximaOrdem(g_empresa, xNumeroEcf, xNumeroCupom, xData, xOrdem) = False
                xValor = xValor + MovCupomFiscal.ValorTotal
                xOrdem = xOrdem + 1
            Loop
            If MovCupomFiscal.LocalizarCodigo(g_empresa, xNumeroEcf, xNumeroCupom, xData, xOrdem) Then
                If rstMovCupomFiscal!Quantidade = 0 And rstMovCupomFiscal![Valor Total] = 0 Then
                    MovCupomFiscal.CupomCancelado = True
                    MovCupomFiscal.ItemCancelado = True
                    If MovCupomFiscalItem.LocalizarCodigo(g_empresa, xNumeroEcf, xData, xNumeroCupom, xOrdem) Then
                        MovCupomFiscalItem.ItemCancelado = True
                        If Not MovCupomFiscalItem.Alterar(g_empresa, xNumeroEcf, xData, xNumeroCupom, xOrdem) Then
                            MsgBox "Não foi possível alterar item cancelado para verdadeiro!", vbInformation, "Erro de Integridade"
                        End If
                        If Not MovCupomFiscal.Alterar(g_empresa, xNumeroEcf, xNumeroCupom, xData, xOrdem) Then
                            MsgBox "Não foi possível alterar cupom cancelado para verdadeiro!", vbInformation, "Erro de Integridade"
                        End If
                    End If
                End If
                MovCupomFiscal.FormaPagamento = 1
                MovCupomFiscal.ValorRecebido = xValor
                If Not MovCupomFiscal.AlterarFormaPagamento(g_empresa, xNumeroEcf, xNumeroCupom, xData) Then
                    MsgBox "Não foi possível alterar a forma de pagamento!", vbInformation, "Erro de Integridade"
                End If
            Else
                MsgBox "Não foi localizar o cupom fiscal", vbInformation, "Erro de Integridade"
            End If
            rstMovCupomFiscal.MoveNext
        Loop
        MsgBox "Processamento concluído!", vbInformation, "Atenção!"
    Else
        MsgBox "Não tem cupons com forma de pagamento zerada!", vbInformation, "Atenção!"
    End If
    rstMovCupomFiscal.Close
    Set rstMovCupomFiscal = Nothing
End Sub
Private Sub zzExcluiBaixaDuplicada()
    On Error GoTo trata_erro
    Dim xBico As Integer
    Dim xData As Date
    Dim xHora As Date
    Dim xEncerrante As String
    Dim xHoraBaixa As Date
    Dim xQtd As Long
    Dim xQtdExcluido As Integer
    Dim xStrData As String
    Dim xDataBase As Date
    
    xStrData = InputBox("Informe a Data no formato dd/mm/yyyy.", "Exclui Duplicidade Abastecimento!", "")
    If Not IsDate(xStrData) Then
        Exit Sub
    End If
    xDataBase = CDate(xStrData)
    xQtdExcluido = 0
    
    
    lSQL = "SELECT *"
    lSQL = lSQL & "  FROM BaixaAbastecimento"
    lSQL = lSQL & " WHERE Data = " & preparaData(xDataBase)
    'lSQL = lSQL & "   AND Bico = 7"
    lSQL = lSQL & " ORDER BY Bico, Data, Hora, [Hora da Baixa]"
    Set rstAbastecimento = Conectar.RsConexao(lSQL)
    If rstAbastecimento.RecordCount > 0 Then
        Do Until rstAbastecimento.EOF
            If rstAbastecimento("Bico").Value <> xBico Or rstAbastecimento("Data").Value <> xData Or rstAbastecimento("Hora").Value <> xHora Or rstAbastecimento("Encerrante").Value <> xEncerrante Then
                xBico = rstAbastecimento("Bico").Value
                xData = rstAbastecimento("Data").Value
                xHora = rstAbastecimento("Hora").Value
                xEncerrante = rstAbastecimento("Encerrante").Value
                xHoraBaixa = rstAbastecimento("Hora da Baixa").Value
            Else
                lSQL = "DELETE BaixaAbastecimento"
                lSQL = lSQL & " WHERE Data = " & preparaData(rstAbastecimento("Data").Value)
                lSQL = lSQL & "   AND Bico = " & rstAbastecimento("Bico").Value
                lSQL = lSQL & "   AND Hora = " & preparaHora(rstAbastecimento("Hora").Value)
                lSQL = lSQL & "   AND Encerrante = " & preparaTexto(rstAbastecimento("Encerrante").Value)
                lSQL = lSQL & "   AND [Hora da Baixa] = " & preparaHora(rstAbastecimento("Hora da Baixa").Value)
                xQtd = Conectar.ExecutaSql(lSQL)
                xQtdExcluido = xQtdExcluido + xQtd
            End If
            rstAbastecimento.MoveNext
        Loop
    End If
    MsgBox "Foi excluido " & xQtdExcluido & " registros."
    rstAbastecimento.Close
    Exit Sub

trata_erro:
    Call CriaLogCupom("Erro zzExcluiBaixaDuplicada: Erro=" & Err.Number & " - " & Err.Description)
End Sub

Private Sub Finaliza()
    'Call TestaConexao(90, "Finaliza")
    Call GravaAuditoria(1, Me.name, 11, "")
    If lNotificacaoGic Then
        menu_personalizado.AtivaVerificacaoGIC
    End If
    Set AberturaCaixa = Nothing
    Set Aliquota = Nothing
    Set BaixaAbastecimento = Nothing
    Set Bomba = Nothing
    Set CartaoAbastecimento = Nothing
    Set CartaoCredito = Nothing
    Set CidadeIBGE = Nothing
    Set Cliente = Nothing
    Set Combustivel = Nothing
    Set Configuracao = Nothing
    Set ConfiguracaoDiversa = Nothing
    Set CerradoTef = Nothing
    Set Credito = Nothing
    Set DuplicataReceber = Nothing
    Set ECF = Nothing
    Set Estoque = Nothing
    Set EncerranteAtual = Nothing
    Set Funcionario = Nothing
    Set Grupo = Nothing
    Set GrupoTipoMovimentoCaixa = Nothing
    Set IntegracaoCaixa = Nothing
    Set LiberacaoDigitacao = Nothing
    Set PeriodoTrocaOleo = Nothing
    Set Produto = Nothing
    Set Produto2 = Nothing
    Set MovCaixaPista = Nothing
    Set MovCartaoCredito = Nothing
    Set MovCupomFiscal = Nothing
    Set MovCupomFiscalItem = Nothing
    Set MovDescontoGrupoCliente = Nothing
    Set MovDescontoPersonalizado = Nothing
    Set MovHorarioVerao = Nothing
    Set MovMapaResumo = Nothing
    Set MovNotaAbastecimento = Nothing
    Set MovimentoAbastecimento = Nothing
    Set MovimentoBomba = Nothing
    Set MovimentoBombaEscritorio = Nothing
    Set MovimentoCheque = Nothing
    Set MovimentoChequeDevolvido = Nothing
    Set PercentualImposto = Nothing
    Set SubEstoque = Nothing
    Set SolicitacaoFuncaoAutomacao = Nothing
    Set TaxaAdmCartaoCredito = Nothing
    Set TicketCarDePara = Nothing
    Set Usuario = Nothing
    'Call CriaLogCupom("Bematech_FI_FechaPortaSerial")
    'BemaRetorno = Bematech_FI_FechaPortaSerial()
    'Call CriaLogCupom("Bematech_FI_FechaPortaSerial - BemaRetorno=" & BemaRetorno)
End Sub
Private Sub GeraDescontoAbastecimento(ByVal pBico As Integer, ByVal pDataAutomacao As Date, ByVal pHoraAutomacao As Date, ByVal pValorLitro As Currency, ByVal pTotalBruto As Currency)
    Dim xValor As String
    Dim xValorDesconto As Currency
    Dim xNovaQuantidade As Currency
    Dim xComplementoDadosInterno As String
    Dim xComplementoCaixa As String
    
    xValor = Format(pTotalBruto, "00000000.00")
    If Val(Mid(xValor, 10, 2)) > 0 And Val(Mid(xValor, 10, 2)) <= 20 Then
        xValorDesconto = Val(Mid(xValor, 10, 2)) / 100
        If MsgBox("Deseja realmente gerar desconto automático?", vbYesNo + vbQuestion + vbDefaultButton2, "Desconto Automático!") = vbYes Then
            xNovaQuantidade = Format((pTotalBruto - xValorDesconto) / pValorLitro, "00000000.000")
            xComplementoCaixa = "Desconto. De: " & Format(pTotalBruto, "###,##0.00") & " Para: " & Format(pTotalBruto - xValorDesconto, "###,##0.00")
            xComplementoDadosInterno = Format(pDataAutomacao, "dd/MM/yyyy") & "|@|" & Format(pHoraAutomacao, "HH:mm:SS") & "|@|" & pBico & "|@|"
            Call GravaAuditoria(1, Me.name, 31, "Func.:" & l_nome_funcionario)
            Call GravaAuditoria(1, Me.name, 31, xComplementoCaixa)
            Call GravaAuditoria(1, Me.name, 31, "Data Aut:" & Format(pDataAutomacao, "dd/MM/yyyy") & " AS " & Format(pHoraAutomacao, "HH:mm:SS") & " Bico:" & pBico)
            If MovimentoAbastecimento.AlterarDesconto(g_empresa, pDataAutomacao, pHoraAutomacao, pBico, xValorDesconto, xNovaQuantidade) Then
                If Not IncluiMovimentoCaixa(lDataCupom, lPeriodo, False, "DESCONTO AUTORIZADO", xValorDesconto, xComplementoDadosInterno, xComplementoCaixa) Then
                    MsgBox "O desconto concedido não foi integrado no caixa!", vbInformation, "Erro de Integridade"
                End If
                AtualizaBombasAbastecimento
            Else
                MsgBox "Não foi possível gerar desconto automático!", vbCritical, "Erro de Integridade!"
            End If
        End If
    End If
End Sub
Private Sub PosicionamentoInicialBombas()
    'MsgBox "PosicionamentoInicialBombas iniciado- TESTE ALEX"
    Dim i As Integer
    Dim xLeft As Integer
    Dim xTop As Integer
    
    For i = 0 To gQUANTIDADE_MAXIMA_BICO - 1
        cmd_bico(i).Visible = False
        cmd_bico(i).Enabled = False
        lbl_automacao_valor(i).Visible = False
        lbl_automacao_valor(i).Enabled = False
        lAutomacaoStatusBico(i) = 0
        lAutomacaoCodigoProduto(i) = 0
        lAutomacaoBico(i) = 0
        lAutomacaoData(i) = CDate("00:00:00")
        lAutomacaoHora(i) = CDate("00:00:00")
        lAutomacaoTempoAbastecimento(i) = ""
        lAutomacaoValorLitro(i) = 0
        lAutomacaoLitros(i) = 0
        lAutomacaoTotalAPagar(i) = 0
    Next
    For i = 0 To (lQtdBomba - 1)
        cmd_bico(i).Visible = True
        cmd_bico(i).Enabled = True
        lbl_automacao_valor(i).Visible = True
        lbl_automacao_valor(i).Enabled = True
    Next
    
    If lQtdBomba = 8 Then
        xTop = 1600
        xLeft = -800
        For i = 0 To 3
            xLeft = xLeft + 1300
            cmd_bico(i).Top = xTop
            cmd_bico(i).Left = xLeft
            lbl_automacao_valor(i).Top = xTop + 750
            lbl_automacao_valor(i).Left = xLeft
        Next
        xTop = 3400
        xLeft = -800
        For i = 4 To (lQtdBomba - 1)
            xLeft = xLeft + 1300
            cmd_bico(i).Top = xTop
            cmd_bico(i).Left = xLeft
            lbl_automacao_valor(i).Top = xTop + 750
            lbl_automacao_valor(i).Left = xLeft
        Next
    End If
    'MsgBox "PosicionamentoInicialBombas finalizado- TESTE ALEX"

End Sub

Private Sub PreencheDicionarioCSTPisCofins()

    If CSTPisValidos.Count > 0 And CSTCofinsValidos.Count > 0 Then Exit Sub


    Dim rs As New ADODB.Recordset


        lSQL = "SELECT CstCofins_TributacaoPisCofins, CstPis_TributacaoPisCofins, AliquotaCofins_TributacaoPisCofins, AliquotaPis_TributacaoPisCofins "
        lSQL = lSQL & " FROM TributacaoPisCofins "
        'lSQL = lSQL & " WHERE IdEstabelecimento_TributacaoPisCofins = 1"
        lSQL = lSQL & " WHERE Saida_TributacaoPisCofins = " & preparaBooleano(True)
        
        Set rs = Conectar.RsConexao(lSQL)
        
        With rs
            If .RecordCount > 0 Then
                .MoveFirst
                Do Until .EOF
                    ''key = CST | Item = aliquota
                     If CSTCofinsValidos.Exists(Val(rs("CstCofins_TributacaoPisCofins").Value)) = False Then
                        Call CSTCofinsValidos.Add(Val(rs("CstCofins_TributacaoPisCofins").Value), CDec(rs("AliquotaCofins_TributacaoPisCofins").Value))
                     End If
                     If CSTPisValidos.Exists(Val(rs("CstPis_TributacaoPisCofins").Value)) = False Then
                        Call CSTPisValidos.Add(Val(rs("CstPis_TributacaoPisCofins").Value), CDec(rs("AliquotaPis_TributacaoPisCofins").Value))
                     End If
                    .MoveNext
                Loop
            End If
            .Close
        End With
        Set rs = Nothing
        
''key = CST | Item = se é tributado
'
'    If CSTPisCofinsValidos.Count > 0 Then Exit Sub
'
'    Call CSTPisCofinsValidos.Add(1, True)
'    Call CSTPisCofinsValidos.Add(2, True)
'    Call CSTPisCofinsValidos.Add(4, False)
   

End Sub

Private Sub PreencheCboBico()
    Dim i As Integer
    cboBico.Clear
    For i = 1 To lQtdBomba
        cboBico.AddItem i
        cboBico.ItemData(cboBico.NewIndex) = i
    Next
End Sub
Private Sub PreencheCboFormaPagamento()
    cbo_forma_pagamento.Clear
    cbo_forma_pagamento.AddItem "1 - Dinheiro"
    cbo_forma_pagamento.ItemData(cbo_forma_pagamento.NewIndex) = 1
    cbo_forma_pagamento.AddItem "2 - Cheque à Vista"
    cbo_forma_pagamento.ItemData(cbo_forma_pagamento.NewIndex) = 2
    cbo_forma_pagamento.AddItem "3 - Cheque Pré-Datado"
    cbo_forma_pagamento.ItemData(cbo_forma_pagamento.NewIndex) = 3
    cbo_forma_pagamento.AddItem "4 - Cartão de Crédito"
    cbo_forma_pagamento.ItemData(cbo_forma_pagamento.NewIndex) = 4
    cbo_forma_pagamento.AddItem "5 - Nota Vinculada"
    cbo_forma_pagamento.ItemData(cbo_forma_pagamento.NewIndex) = 5
    cbo_forma_pagamento.AddItem "6 - Cartão TecBan    "
    cbo_forma_pagamento.ItemData(cbo_forma_pagamento.NewIndex) = 6
    cbo_forma_pagamento.AddItem "7 - Cheque TecBan    "
    cbo_forma_pagamento.ItemData(cbo_forma_pagamento.NewIndex) = 7
    cbo_forma_pagamento.AddItem "8 - Ticket Car Smart "
    cbo_forma_pagamento.ItemData(cbo_forma_pagamento.NewIndex) = 8
    cbo_forma_pagamento.AddItem "9 - Smart Shop/Check Check"
    cbo_forma_pagamento.ItemData(cbo_forma_pagamento.NewIndex) = 9
    cbo_forma_pagamento.AddItem "10 - SuperCard"
    cbo_forma_pagamento.ItemData(cbo_forma_pagamento.NewIndex) = 10
    cbo_forma_pagamento.AddItem "11 - HiperCard"
    cbo_forma_pagamento.ItemData(cbo_forma_pagamento.NewIndex) = 11
    cbo_forma_pagamento.AddItem "12 - PagCard"
    cbo_forma_pagamento.ItemData(cbo_forma_pagamento.NewIndex) = 12
    cbo_forma_pagamento.AddItem "13 - Cheque Redecard  "
    cbo_forma_pagamento.ItemData(cbo_forma_pagamento.NewIndex) = 13
    cbo_forma_pagamento.AddItem "14 - USA Card"
    cbo_forma_pagamento.ItemData(cbo_forma_pagamento.NewIndex) = 14
    cbo_forma_pagamento.AddItem "15 - GodCard"
    cbo_forma_pagamento.ItemData(cbo_forma_pagamento.NewIndex) = 15
    cbo_forma_pagamento.AddItem "16 - Cartão pelo POS"
    cbo_forma_pagamento.ItemData(cbo_forma_pagamento.NewIndex) = 16
    cbo_forma_pagamento.AddItem "17 - Cerrado Tef"
    cbo_forma_pagamento.ItemData(cbo_forma_pagamento.NewIndex) = 17

End Sub

Private Sub PreencheCboFormaPagamentoParaPostoAki()
    cbo_forma_pagamento.Clear
    cbo_forma_pagamento.AddItem "1 - Dinheiro"
    cbo_forma_pagamento.ItemData(cbo_forma_pagamento.NewIndex) = 1
    cbo_forma_pagamento.AddItem "4 - Cartão de Crédito"
    cbo_forma_pagamento.ItemData(cbo_forma_pagamento.NewIndex) = 4
    cbo_forma_pagamento.AddItem "16 - Cartão pelo POS"
    cbo_forma_pagamento.ItemData(cbo_forma_pagamento.NewIndex) = 16
    cbo_forma_pagamento.ListIndex = 0
    cbo_forma_pagamento.SetFocus
End Sub


Private Sub PreencheCboTipoSubEstoque()
    Dim rstTipoSubEstoque As ADODB.Recordset
    
    cboTipoSubEstoque.Clear
    Set rstTipoSubEstoque = Conectar.RsConexao("SELECT Codigo, Nome FROM TipoSubEstoque WHERE Codigo > 1 ORDER BY Codigo")
    Do Until rstTipoSubEstoque.EOF
        cboTipoSubEstoque.AddItem rstTipoSubEstoque!Codigo & " " & rstTipoSubEstoque!Nome
        cboTipoSubEstoque.ItemData(cboTipoSubEstoque.NewIndex) = rstTipoSubEstoque!Codigo
        rstTipoSubEstoque.MoveNext
    Loop
    rstTipoSubEstoque.Close
    Set rstTipoSubEstoque = Nothing
End Sub
Private Sub PreencheDadosDocumentoEletronicoCabecalho() 'ALEX - DOCELETRONICO
    
    On Error GoTo trata_erro
    lNumeroUltimoCupom = lNumeroCupom
    
    'Call PreparaTipoMovimento(Produto.CodigoGrupo)
    
    MovDocEletronicoCabecalho.IdEstabelecimento = g_empresa
    MovDocEletronicoCabecalho.DataEmissao = lData
    MovDocEletronicoCabecalho.Entrada = False
    MovDocEletronicoCabecalho.Saida = True
    MovDocEletronicoCabecalho.Modelo = MODELO_NFCE
    MovDocEletronicoCabecalho.Serie = lSerieNFCe
    MovDocEletronicoCabecalho.numero = lNumeroCupom
    MovDocEletronicoCabecalho.HoraSaida = lHora
    MovDocEletronicoCabecalho.DataEntradaSaida = lData
    MovDocEletronicoCabecalho.EmissaoPropria = True
    MovDocEletronicoCabecalho.IdClienteFornecedor = l_codigo_cliente
    MovDocEletronicoCabecalho.CodigoSituacao = 0
    MovDocEletronicoCabecalho.ChaveAcesso = "0"
    MovDocEletronicoCabecalho.FormaPagamento = "1"
    MovDocEletronicoCabecalho.ValorTotal = fValidaValor2(txt_valor_total.Text)
    MovDocEletronicoCabecalho.ValorDesconto = fValidaValor2(txt_valor_desconto.Text)
    MovDocEletronicoCabecalho.ValorAbatimentoNaoTributado = 0
    MovDocEletronicoCabecalho.ValorProdutos = fValidaValor2(txt_valor_total.Text)
    MovDocEletronicoCabecalho.TipoFrete = "1"
    MovDocEletronicoCabecalho.ValorFrete = 0
    MovDocEletronicoCabecalho.ValorSeguro = 0
    MovDocEletronicoCabecalho.OutrasDespesas = 0
    MovDocEletronicoCabecalho.ValorBCICMS = IIf(Aliquota.Aliquota > 0, fValidaValor2(txt_valor_total.Text), 0)  'fValidaValor2(txt_valor_total.Text)
    MovDocEletronicoCabecalho.AliquotaICMS = 0
    MovDocEletronicoCabecalho.ValorICMS = 0
    MovDocEletronicoCabecalho.ValorBCICMSST = 0 'IIf(Aliquota.Aliquota > 0, fValidaValor2(txt_valor_total.Text), 0)  'fValidaValor2(txt_valor_total.Text)
    MovDocEletronicoCabecalho.AliquotaIcmsSt = 0
    MovDocEletronicoCabecalho.ValorICMSST = 0
    MovDocEletronicoCabecalho.ValorIPI = 0
    MovDocEletronicoCabecalho.ValorPis = 0
    MovDocEletronicoCabecalho.ValorCofins = 0
    MovDocEletronicoCabecalho.ValorPisSt = 0
    MovDocEletronicoCabecalho.ValorCofinsSt = 0
    MovDocEletronicoCabecalho.Combustivel = IIf(Produto.TipoCombustivel = Empty, False, True)
    MovDocEletronicoCabecalho.Cancelado = False
    MovDocEletronicoCabecalho.AguaEnergiaGasTelefone = ""
    MovDocEletronicoCabecalho.Inutilizada = False
    MovDocEletronicoCabecalho.IncidePisConfis = False
    MovDocEletronicoCabecalho.DataDigitacao = lData
    MovDocEletronicoCabecalho.DataAlteracao = "00:00:00"
    MovDocEletronicoCabecalho.IdUsuario = l_codigo_funcionario
    MovDocEletronicoCabecalho.NumeroLote = 0
    MovDocEletronicoCabecalho.NumeroRecepcao = 0
    MovDocEletronicoCabecalho.NumeroProtocolo = 0
    MovDocEletronicoCabecalho.EtapaConcluida = Val(ETAPA_CONCLUIDA.GRAVADO)
    MovDocEletronicoCabecalho.NomeEtapaConcluida = "GRAVADO"
    MovDocEletronicoCabecalho.CodigoUltimoEvento = Val(EVENTO_NFCE.NENHUM_EVENTO)
    MovDocEletronicoCabecalho.ObservacaoEvento = ""
    MovDocEletronicoCabecalho.ProgramaOrigem = PROGRAMA_ORIGEM
    MovDocEletronicoCabecalho.Periodo = lPeriodo
    MovDocEletronicoCabecalho.TipoMovimento = lTipoMovimento
    MovDocEletronicoCabecalho.TipoSubEstoque = Val(cboTipoSubEstoque.Text)
    
    Exit Sub

trata_erro:
    Call CriaLogCupom("Erro PreencheDadosDocumentoEletronicoCabecalho: Erro=" & Err.Number & " - " & Err.Description)
End Sub
Private Sub PreparaDadosAdicionaisFechamento()
    lbl_numero_cheque.Visible = True
    txt_numero_cheque.Visible = True
    lbl_telefone.Visible = True
    txt_telefone.Visible = True
    lbl_valor_recebido.Left = 3180
    txt_valor_recebido.Left = 3180
    lbl_valor_troco1.Left = 4620
    lbl_valor_troco.Left = 4620
    cmd_cancelar2.Left = 3800
    cmdIniciaProcessoNFCe.Left = 4820
    cmd_cancelar2.Top = 3900
    cmdIniciaProcessoNFCe.Top = 3900
    'frm_fechamento_cupom.Top = 3300
    'frm_fechamento_cupom.Left = 5970
    'frm_fechamento_cupom.Height = 4350
    'frm_fechamento_cupom.Width = 5800
End Sub
Private Function PreparaDadosProdutos() As String
    Dim rsProdutosECF As ADODB.Recordset
    Dim xSQL As String
    Dim xString As String
    
    PreparaDadosProdutos = ""
    xString = ""
    xSQL = ""
    xSQL = xSQL & "SELECT Movimento_Cupom_Fiscal.[Codigo do Produto],"
    xSQL = xSQL & "       Movimento_Cupom_Fiscal.Quantidade,"
    xSQL = xSQL & "       Movimento_Cupom_Fiscal.[Valor Total],"
    xSQL = xSQL & "       Movimento_Cupom_Fiscal.[Tipo de Combustivel],"
    xSQL = xSQL & "       Produto.Nome"
    xSQL = xSQL & "  FROM Movimento_Cupom_Fiscal, Produto"
    xSQL = xSQL & " WHERE Movimento_Cupom_Fiscal.Empresa = " & g_empresa
    xSQL = xSQL & "   AND Movimento_Cupom_Fiscal.[Codigo da ECF] = " & lNumeroPDV
    xSQL = xSQL & "   AND Movimento_Cupom_Fiscal.[Numero do Cupom] = " & lNumeroCupom
    xSQL = xSQL & "   AND Movimento_Cupom_Fiscal.Data = " & preparaData(lData)
    xSQL = xSQL & "   AND Movimento_Cupom_Fiscal.[Cupom Cancelado] = " & preparaBooleano(False)
    xSQL = xSQL & "   AND Movimento_Cupom_Fiscal.[Item Cancelado] = " & preparaBooleano(False)
    xSQL = xSQL & "   AND Movimento_Cupom_Fiscal.[Codigo do Produto] = Produto.Codigo"
    xSQL = xSQL & " ORDER BY Produto.Nome, Movimento_Cupom_Fiscal.Ordem"
    Set rsProdutosECF = Conectar.RsConexao(xSQL)
    Do Until rsProdutosECF.EOF
        xString = xString & rsProdutosECF("Codigo do Produto").Value & "|@|"
        xString = xString & rsProdutosECF("Nome").Value & "|@|"
        xString = xString & rsProdutosECF("Quantidade").Value & "|@|"
        xString = xString & rsProdutosECF("Valor Total").Value & "|@|"
        xString = xString & rsProdutosECF("Tipo de Combustivel").Value & "|@|" & vbCrLf
        rsProdutosECF.MoveNext
    Loop
    rsProdutosECF.Close
    Set rsProdutosECF = Nothing
    PreparaDadosProdutos = xString
End Function
Private Sub PreparaTipoMovimento(ByVal pCodigoGrupo As Integer)
    Dim xTipoVenda As String
    
    xTipoVenda = ReadINI("CUPOM FISCAL", "Tipo de Venda", gArquivoIni)
    If xTipoVenda = "AUTOMACAO/CONVENIENCIA" Then
        lTipoMovimento = 1
    Else
        lTipoMovimento = 2
    End If
    If pCodigoGrupo > 0 Then
        If GrupoTipoMovimentoCaixa.LocalizarGrupo(pCodigoGrupo) Then
            lTipoMovimento = GrupoTipoMovimentoCaixa.TipoMovimento
        End If
    End If
    If PeriodoTrocaOleo.LocalizarCodigo(g_empresa, Val(txt_funcionario_ponto.Text)) Then
        lTipoMovimento = 3
        cboTipoSubEstoque.ListIndex = lTipoMovimento - 2
    End If
End Sub
Private Function VerificaClienteEmAtraso() As Boolean
    Dim xTolerancia As Integer
    
    xTolerancia = 1
    If DuplicataReceber.LocalizaPrimeiroVencCliente(Cliente.Codigo) Then
        If Credito.LocalizarCodigo(Cliente.Codigo) Then
            xTolerancia = Credito.DiasAtraso
        End If
        If DateDiff("d", DuplicataReceber.DataVencimento, Date) > xTolerancia Then
            MsgBox "Cliente tem duplicata em aberto:" & Chr(10) & "Vencimento: " & Format(DuplicataReceber.DataVencimento, "dd/mm/yyyy") & Chr(10) & "Valor: " & Format(DuplicataReceber.ValorVencimento & Chr(10) & "Com " & DateDiff("d", DuplicataReceber.DataVencimento, Date) & " dias de atraso." & Chr(10) & "A tolerância era de " & xTolerancia & " dias de atraso.", "###,###,##0.00"), vbInformation, "Duplicata em Aberto"
        End If
    End If
End Function
Private Function VerificaSaldoCreditoCliente(ByVal pValorVenda As Currency) As Boolean
    Dim xValorCredito As Currency
    Dim xValorDuplicata As Currency
    Dim xValorNotaAbastecimento As Currency
    Dim xValorItemImprimir As Currency
    Dim xValorCupomAtual As Currency
    Dim xValorSomatorioDebitos As Currency
    Dim xValorCreditoRestante As Currency
    Dim xValorCupomJaImpresso As Currency
    
    
    VerificaSaldoCreditoCliente = False
    xValorCredito = 0
    xValorDuplicata = 0
    xValorNotaAbastecimento = 0
    xValorCupomJaImpresso = 0
    
    If Credito.LocalizarCodigo(Cliente.Codigo) Then
        xValorCredito = Credito.Limite
    End If
    
    xValorDuplicata = DuplicataReceber.TotalCliente(g_empresa, Cliente.Codigo)
    xValorNotaAbastecimento = MovNotaAbastecimento.TotalCliente(g_empresa, Cliente.Codigo)
    xValorItemImprimir = fValidaValor(txt_valor_total.Text)
    
    xValorCupomJaImpresso = 0
    If lNumeroCupom = lNumeroUltimoCupom Then
        xValorCupomJaImpresso = lTotalCupom
    End If
    xValorCupomAtual = xValorCupomJaImpresso + xValorItemImprimir
    xValorSomatorioDebitos = xValorNotaAbastecimento + xValorDuplicata + xValorCupomAtual
    xValorCreditoRestante = xValorCredito - (xValorNotaAbastecimento + xValorDuplicata + xValorCupomJaImpresso)
    
    If (pValorVenda + xValorDuplicata + xValorNotaAbastecimento) > xValorCredito Then
        Dim xMensagem As String
        xMensagem = "O Valor das somas dos débitos é maior que o crédito do cliente." & vbCrLf & vbCrLf
        xMensagem = xMensagem & "Valor do Item a Imprimir: " & Format(xValorItemImprimir, "###,###,##0.00") & vbCrLf
        xMensagem = xMensagem & "Valor do Cupom já Impresso: " & Format(xValorCupomJaImpresso, "###,###,##0.00") & vbCrLf
        xMensagem = xMensagem & "Soma dos Valores do Cupom Atual: " & Format(xValorCupomAtual, "###,###,##0.00") & vbCrLf & vbCrLf
        xMensagem = xMensagem & "Valor das Duplicatas: " & Format(xValorDuplicata, "###,###,##0.00") & vbCrLf
        xMensagem = xMensagem & "Valor das Notas de Abastecimento: " & Format(xValorNotaAbastecimento, "###,###,##0.00") & vbCrLf
        xMensagem = xMensagem & "Somatório: " & Format(xValorSomatorioDebitos, "###,###,##0.00") & vbCrLf & vbCrLf
        xMensagem = xMensagem & "Saldo do Cliente Concedido no Cadastro: " & Format(xValorCredito, "###,###,##0.00") & vbCrLf
        xMensagem = xMensagem & "Saldo de Crédito Restante: " & Format(xValorCreditoRestante, "###,###,##0.00") & vbCrLf
        MsgBox xMensagem & Chr(10), vbInformation, "Crédito Insuficiente!"
    Else
        VerificaSaldoCreditoCliente = True
    End If
    
End Function
Function VerificaDataHora() As Boolean
    Dim xData As Date
    Dim xHora As Date
    
    On Error GoTo FileError
    
    VerificaDataHora = False
    xData = ReadINI("CUPOM FISCAL", "Informa Encerrante na Data", gArquivoIni)
    xHora = ReadINI("CUPOM FISCAL", "Informa Encerrante na Hora", gArquivoIni)
    If Date > xData Then
        xData = Date
        xHora = CDate("01:00:00")
    End If
    If Date >= xData Then
        If Time >= xHora Then
            g_string = "deletar" & "|@|" & xData & "|@|"
            movimento_bomba.Show 1
            Exit Function
        End If
    End If
    Exit Function

FileError:
    Call CriaLogCupom("Erro VerificaDataHora: Erro=" & Err.Number & " - " & Err.Description)
    If Err = 62 Then
        Exit Function
    End If
    MsgBox "O arquivo temporizador DATAHORA.TXT não foi encontrado!", vbCritical, "Arquivo não Encontrado!"
    Exit Function
End Function
Private Sub VerificaLiberacaoDigitacao()
    Dim x_flag As Boolean
    x_flag = True
    If g_nivel_acesso > 4 Then
        If MovCupomFiscal.Empresa < g_cfg_empresa_i Or MovCupomFiscal.Empresa > g_cfg_empresa_f Then
            x_flag = False
        ElseIf MovCupomFiscal.Data < g_cfg_data_i Or MovCupomFiscal.Data > g_cfg_data_f Then
            x_flag = False
        ElseIf MovCupomFiscal.Periodo < g_cfg_periodo_i Or MovCupomFiscal.Periodo > g_cfg_periodo_f Then
            x_flag = False
        End If
    End If
End Sub
Function VerificaLiberacaoDigitacao2() As Boolean
    Dim xTipoMovimento As Integer
    
    xTipoMovimento = lTipoMovimento
    If cboTipoSubEstoque.ListCount > 1 And UCase(Funcionario.Cargo) Like "*TROCADOR*" Then
        xTipoMovimento = 3
    End If
    VerificaLiberacaoDigitacao2 = False
    
    If lCaixaIndividual Then
        If Not AberturaCaixa.LocalizarUltAbertoDataFunc(g_empresa, lData, "NF", 1, xTipoMovimento, l_codigo_funcionario) Then
            If (MsgBox("O Caixa não encontra-se aberto!" & Chr(10) & "Deseja abrir agora?", vbQuestion + vbYesNo + vbDefaultButton1, "Abertura de Caixa!")) = vbYes Then
                If CriaAberturaCaixa(xTipoMovimento, lPeriodo) = False Then
                    Exit Function
                End If
            Else
                Exit Function
            End If
        End If
    Else
        If Not AberturaCaixa.LocalizarCxData(g_empresa, lData, "NF", lPeriodo, 1, xTipoMovimento) Then
            If (MsgBox("O Caixa não encontra-se aberto!" & Chr(10) & "Deseja abrir agora?", vbQuestion + vbYesNo + vbDefaultButton1, "Abertura de Caixa!")) = vbYes Then
                If CriaAberturaCaixa(xTipoMovimento, lPeriodo) = False Then
                    Exit Function
                End If
            Else
                Exit Function
            End If
        End If
    End If
    If g_nivel_acesso <= 4 Then
        VerificaLiberacaoDigitacao2 = True
        Exit Function
    End If
    If lData < g_cfg_data_i Or lData > g_cfg_data_f Then
        MsgBox "A data do cupom deve estar entre " & Format(g_cfg_data_i, "dd/mm/yyyy") & " a " & Format(g_cfg_data_f, "dd/mm/yyyy") & ".", vbInformation, "Digitação Não Autorizada!"
        txt_produto.SetFocus
    ElseIf lPeriodo < g_cfg_periodo_i Or lPeriodo > g_cfg_periodo_f Then
        MsgBox "O período deve estar entre " & g_cfg_periodo_i & " ao " & g_cfg_periodo_f & ".", vbInformation, "Digitação Não Autorizada!"
        txt_produto.SetFocus
    Else
        VerificaLiberacaoDigitacao2 = True
    End If
End Function
Private Function MontaTextoInformacoesComplementaresNFCe() As String

    Dim xTextoInformacoesComplementares As String
    Dim xLinhaImpostos As String
    Dim xUtilizaTecnoSpeed As Boolean
    Dim xQuebraLinhaImpressao As String
    
    xUtilizaTecnoSpeed = False
    xTextoInformacoesComplementares = ""
    xQuebraLinhaImpressao = "\n"
    
    xLinhaImpostos = CalculaImpostos(lNumeroCupom, lData)
        
    If ConfiguracaoDiversa.LocalizarCodigo(g_empresa, "NFCe Imprimir Atraves") Then
        If ConfiguracaoDiversa.Texto = "TECNOSPEED" Then
            xQuebraLinhaImpressao = "|"
        End If
    End If

    xTextoInformacoesComplementares = xTextoInformacoesComplementares & "Funcionario: " + Format(l_codigo_funcionario, "000") + " " + Mid(l_nome_funcionario, 1, 30) + xQuebraLinhaImpressao
    
    If Not Trim(xLinhaImpostos) = Empty Then
        xTextoInformacoesComplementares = xTextoInformacoesComplementares & xLinhaImpostos & xQuebraLinhaImpressao
    End If
    
    If l_codigo_cliente = 0 Then
        If Not Trim(txt_observacao.Text) = Empty Then
            xTextoInformacoesComplementares = xTextoInformacoesComplementares & txt_observacao.Text & xQuebraLinhaImpressao
        End If
        
        If Not Trim(txt_observacao_2.Text) = Empty Then
            xTextoInformacoesComplementares = xTextoInformacoesComplementares & txt_observacao_2.Text & xQuebraLinhaImpressao
        End If
    End If
    
    If Not Trim(txt_placa.Text) = Empty Then
        xTextoInformacoesComplementares = xTextoInformacoesComplementares & "PLACA: " & txt_placa.Text & xQuebraLinhaImpressao
    End If

    If Not Trim(txt_kilometragem.Text) = Empty Then
        xTextoInformacoesComplementares = xTextoInformacoesComplementares & "KM: " & txt_kilometragem.Text
    End If
    
    If lDescontoPostoAki = True And lValorDescontoConcedido > 0 Then 'ALEX
        xTextoInformacoesComplementares = xTextoInformacoesComplementares & "Desconto pelo Aplicativo PostoAki" & xQuebraLinhaImpressao
    End If
    
    
    MontaTextoInformacoesComplementaresNFCe = xTextoInformacoesComplementares

End Function
Private Function MontaTextoItensSolicitacaoNFCE(ByVal pRsDadosParaNFCe As ADODB.Recordset) As String

    MontaTextoItensSolicitacaoNFCE = Empty
    Dim xOrdem As Integer
    Dim xStringNfce As String
     
    xOrdem = 0
    CriaLogCupom ("[MontaTextoItensSolicitacaoNFCE] - QUANTIDADE DE ITENS: " & pRsDadosParaNFCe.RecordCount)
           Do Until pRsDadosParaNFCe.EOF
           
                xOrdem = xOrdem + 1
                    xStringNfce = xStringNfce & "800-" & Format(xOrdem, "000") & " = INICIO" & "|@|" & vbCrLf
                    xStringNfce = xStringNfce & "810-" & Format(xOrdem, "000") & " = " & pRsDadosParaNFCe("Codigo NCM").Value & "|@|" & vbCrLf
                    xStringNfce = xStringNfce & "811-" & Format(xOrdem, "000") & " = " & pRsDadosParaNFCe("NomeProduto").Value & "|@|" & vbCrLf
                    xStringNfce = xStringNfce & "812-" & Format(xOrdem, "000") & " = " & pRsDadosParaNFCe("Unidade").Value & "|@|" & vbCrLf
                    xStringNfce = xStringNfce & "813-" & Format(xOrdem, "000") & " = " & pRsDadosParaNFCe("IdProduto_MovDEItem").Value & "|@|" & vbCrLf
                    xStringNfce = xStringNfce & "820-" & Format(xOrdem, "000") & " = " & FormatNumber(pRsDadosParaNFCe("ValorUnitario_MovDEItem").Value, 4) & "|@|" & vbCrLf
                    xStringNfce = xStringNfce & "821-" & Format(xOrdem, "000") & " = " & FormatNumber(pRsDadosParaNFCe("Quantidade_MovDEItem").Value, 4) & "|@|" & vbCrLf
                    'xStringNfce = xStringNfce & "822-" & Format(xOrdem, "000") & " = " & FormatNumber(pRsDadosParaNFCe("ValorTotalLiquido_MovDEItem").Value, 2) & "|@|" & vbCrLf
                    xStringNfce = xStringNfce & "822-" & Format(xOrdem, "000") & " = " & FormatNumber(pRsDadosParaNFCe("TotalBruto_MovDEItem").Value, 4) & "|@|" & vbCrLf
                    'xStringNfce = xStringNfce & "822-" & Format(xOrdem, "000") & " = " & "|@|" & vbCrLf '& FormatNumber(dvItensNF(i)("TotalVenda"), 4) & "|@|" & vbCrLf 'TotalProdServ
                    
                    xStringNfce = xStringNfce & "823-" & Format(xOrdem, "000") & " = " & pRsDadosParaNFCe("Codigo de Barra").Value & "|@|" & vbCrLf '& Produto.CodigoBarra & "|@|" & vbCrLf
                    
                    'cst pis
                    xStringNfce = xStringNfce & "824-" & Format(xOrdem, "000") & " = " & pRsDadosParaNFCe("CstPis_MovDEItem").Value & "|@|" & vbCrLf '& Format(Produto.CSTPIS, "00") & "|@|" & vbCrLf
                    'cst cofins
                    xStringNfce = xStringNfce & "825-" & Format(xOrdem, "000") & " = " & pRsDadosParaNFCe("CstCofins_MovDEItem").Value & "|@|" & vbCrLf '& Format(Produto.CSTCOFINS, "00") & "|@|" & vbCrLf
                    
                    xStringNfce = xStringNfce & "826-" & Format(xOrdem, "000") & " = " & "" & "|@|" & vbCrLf 'CodListServico
                    
                    'valor BC pis
                    xStringNfce = xStringNfce & "827-" & Format(xOrdem, "000") & " = " & FormatNumber(pRsDadosParaNFCe("ValorBcPis_MovDEItem").Value, 4) & "|@|" & vbCrLf
                    
                    'valor BC Cofins
                    xStringNfce = xStringNfce & "828-" & Format(xOrdem, "000") & " = " & FormatNumber(pRsDadosParaNFCe("ValorBcCofins_MovDEItem").Value, 4) & "|@|" & vbCrLf
                    xStringNfce = xStringNfce & "829-" & Format(xOrdem, "000") & " = " & "0" & "|@|" & vbCrLf '& lPisValor & "|@|" & vbCrLf 'ValorPisReais
                    xStringNfce = xStringNfce & "830-" & Format(xOrdem, "000") & " = " & "0" & "|@|" & vbCrLf '& lCofinsValor & "|@|" & vbCrLf 'ValorCofinsREais

                    xStringNfce = xStringNfce & "831-" & Format(xOrdem, "000") & " = " & pRsDadosParaNFCe("CFOP_MovDEItem").Value & "|@|" & vbCrLf
                    
                    'Aliquota Pis
                    xStringNfce = xStringNfce & "832-" & Format(xOrdem, "000") & " = " & FormatNumber(pRsDadosParaNFCe("AliquotaPis_MovDEItem").Value, 2) & "|@|" & vbCrLf 'Venda de Combustivel ou lubrificantes" & "|@|" & vbCrLf '& RetiraAcentos(Cfop.NaturezaOperacaoReduzida) & "|@|" & vbCrLf 'Descricao CFOP
                    
                    'Aliquota Cofins
                    xStringNfce = xStringNfce & "833-" & Format(xOrdem, "000") & " = " & FormatNumber(pRsDadosParaNFCe("AliquotaCofins_MovDEItem").Value, 2) & "|@|" & vbCrLf '& xCodigoUfEmitente & "|@|" & vbCrLf 'CodEstadoIde
                    
                    xStringNfce = xStringNfce & "834-" & Format(xOrdem, "000") & " = 0" & "|@|" & vbCrLf '& "0" & "|@|" & vbCrLf 'ValorIPIProd
                    
                    'valor cofins
                    xStringNfce = xStringNfce & "835-" & Format(xOrdem, "000") & " = " & FormatNumber(pRsDadosParaNFCe("ValorCofins_MovDEItem").Value, 2) & "|@|" & vbCrLf '& lCofinsPercentual & "|@|" & vbCrLf 'ValorCofinsProd
                    
                    'valor pis
                    xStringNfce = xStringNfce & "836-" & Format(xOrdem, "000") & " = " & FormatNumber(pRsDadosParaNFCe("ValorPis_MovDEItem").Value, 2) & "|@|" & vbCrLf '& lPisPercentual & "|@|" & vbCrLf 'ValorPISProd
                    
                    'Valor ICMS
                    xStringNfce = xStringNfce & "837-" & Format(xOrdem, "000") & " = " & FormatNumber(pRsDadosParaNFCe("ValorIcms_MovDEItem").Value, 4) & "|@|" & vbCrLf  '& FormatNumber(lIcmsValor, 4) & "|@|" & vbCrLf 'ValorIcmsProd
                    
                    'valor BC icms
                    xStringNfce = xStringNfce & "838-" & Format(xOrdem, "000") & " = " & FormatNumber(pRsDadosParaNFCe("ValorBcIcms_MovDEItem").Value, 4) & "|@|" & vbCrLf  'ValorBCICMSProd
                    
                    xStringNfce = xStringNfce & "839-" & Format(xOrdem, "000") & " = " & "0" & "|@|" & vbCrLf 'AliquotaIPIProd
                    
                    'Aliquota icms
                    xStringNfce = xStringNfce & "840-" & Format(xOrdem, "000") & " = " & FormatNumber(pRsDadosParaNFCe("AliquotaIcms_MovDEItem").Value, 4) & "|@|" & vbCrLf 'FormatNumber(pRsDadosParaNFCe("Aliquota do Imposto").Value, 4) & "|@|" & vbCrLf 'AliquotaICMSProd
                    
                    xStringNfce = xStringNfce & "841-" & Format(xOrdem, "000") & " = " & FormatNumber(pRsDadosParaNFCe("ValorDesconto_MovDEItem").Value, 4) & "|@|" & vbCrLf  '& txtDesconto.Text & "|@|" & vbCrLf 'DescontoProd
                    xStringNfce = xStringNfce & "842-" & Format(xOrdem, "000") & " = " & "0" & "|@|" & vbCrLf 'Descontope
                    xStringNfce = xStringNfce & "843-" & Format(xOrdem, "000") & " = " & "1" & "|@|" & vbCrLf 'tpNF (1-Saída)
                    'cst icms
                    xStringNfce = xStringNfce & "844-" & Format(xOrdem, "000") & " = " & Format(pRsDadosParaNFCe("CstIcms_MovDEItem").Value, "00") & "|@|" & vbCrLf
                    
                   
                    xStringNfce = xStringNfce & "845-" & Format(xOrdem, "000") & " = " & "P" & "|@|" & vbCrLf 'ServProd
                    xStringNfce = xStringNfce & "846-" & Format(xOrdem, "000") & " = " & "|@|" & vbCrLf 'DadosAdicionais
                    
                    'valor bc icms st
                    xStringNfce = xStringNfce & "847-" & Format(xOrdem, "000") & " = " & FormatNumber(pRsDadosParaNFCe("ValorBcIcmsSt_MovDEItem").Value, 4) & "|@|" & vbCrLf 'BCST
                    'valor icms st
                    xStringNfce = xStringNfce & "848-" & Format(xOrdem, "000") & " = " & FormatNumber(pRsDadosParaNFCe("ValorIcmsSt_MovDEItem").Value, 4) & "|@|" & vbCrLf 'ValorIcmsSub
                    'Aliquota icms st
                    xStringNfce = xStringNfce & "849-" & Format(xOrdem, "000") & " = " & FormatNumber(pRsDadosParaNFCe("AliquotaIcmsSt_MovDEItem").Value, 4) & "|@|" & vbCrLf 'PercentualICMSSub
                    
                    xStringNfce = xStringNfce & "850-" & Format(xOrdem, "000") & " = " & "0" & "|@|" & vbCrLf 'PercentualRedICMSSub
                    xStringNfce = xStringNfce & "851-" & Format(xOrdem, "000") & " = " & "0" & "|@|" & vbCrLf 'PercentualRedICMS
                    
                    
                    'Valor do Desconto Total do Ítem
                    xStringNfce = xStringNfce & "853-" & Format(xOrdem, "000") & " = 0" & "|@|" & vbCrLf '& FormatNumber(dvItensNF(i)("Valor do Desconto"), 4) & "|@|" & vbCrLf '
                    'Valor do Acréscimo Total do Ítem
                    xStringNfce = xStringNfce & "854-" & Format(xOrdem, "000") & " = 0" & "|@|" & vbCrLf '& FormatNumber(0, 4) & "|@|" & vbCrLf '

                    'Tipo de Combustivel
                    xStringNfce = xStringNfce & "855-" & Format(xOrdem, "000") & " = " & pRsDadosParaNFCe("TipoCombustivel_MovDEItem").Value & "|@|" & vbCrLf   '& Produto.TipoCombustivel & "|@|" & vbCrLf
                    
                    'Codigo da ANP
                    xStringNfce = xStringNfce & "856-" & Format(xOrdem, "000") & " = " & pRsDadosParaNFCe("Codigo ANP").Value & "|@|" & vbCrLf  '620505001 & Produto.CodigoANP & "|@|" & vbCrLf
                    
                    'Codigo CEST
                    xStringNfce = xStringNfce & "857-" & Format(xOrdem, "000") & " = " & pRsDadosParaNFCe("CEST").Value & "|@|" & vbCrLf '- OBRIGATORIO CASO O TIPO DE TRIBUTAÇÃO SEJA SUBSTITUIÇÃO
                    
                    'ENCERRANTE INICIAL 'CALCULADO PELO ENCERRANTE FINAL - QUANTIDADE
                    Dim xEncerranteFinal As Currency
                    xEncerranteFinal = pRsDadosParaNFCe("EncerranteFinal_MovDEItem").Value
                    xStringNfce = xStringNfce & "858-" & Format(xOrdem, "000") & " = " & RetornaEncerranteInicial(xEncerranteFinal, pRsDadosParaNFCe("Quantidade_MovDEItem").Value) & "|@|" & vbCrLf    '- OBRIGATORIO CASO O TIPO DE TRIBUTAÇÃO SEJA SUBSTITUIÇÃO
                    
                    'ENCERRANTE FINAL 'ESTÁ SENDO GRAVADO NO CAMPO DE TELEFONE DA TABELA MOVIMENTO_CUPOOM_FISCAL
                    xStringNfce = xStringNfce & "859-" & Format(xOrdem, "000") & " = " & xEncerranteFinal & "|@|" & vbCrLf
                    
                    'Bomba ESTÁ SENDO GRAVADO NO CAMPO DE NUMERO DO CHEQUE (POSIÇÕES 1 e 2) DA TABELA MOVIMENTO_CUPOOM_FISCAL - ALTERADO APÓS CRIAÇÃO DA TABELA DE DOCUMENTO ELETRÔNICO
                    xStringNfce = xStringNfce & "860-" & Format(xOrdem, "000") & " = " & pRsDadosParaNFCe("NumeroBomba_MovDEItem").Value & "|@|" & vbCrLf  'Mid(pRsDadosParaNFCe("Numero do Cheque").Value, 1, 2) & "|@|" & vbCrLf
                    
                    'Bico ESTÁ SENDO GRAVADO NO CAMPO DE NUMERO DO CHEQUE (POSIÇÕES 3 e 4) DA TABELA MOVIMENTO_CUPOOM_FISCAL
                    xStringNfce = xStringNfce & "861-" & Format(xOrdem, "000") & " = " & pRsDadosParaNFCe("NumeroBico_MovDEItem").Value & "|@|" & vbCrLf  ' Mid(pRsDadosParaNFCe("Numero do Cheque").Value, 3, 2) & "|@|" & vbCrLf
                    
                    'tanque ESTÁ SENDO GRAVADO NO CAMPO DE NUMERO DO CHEQUE (POSIÇÕES 5 e 6) DA TABELA MOVIMENTO_CUPOOM_FISCAL
                    xStringNfce = xStringNfce & "862-" & Format(xOrdem, "000") & " = " & pRsDadosParaNFCe("NumeroTanque_MovDEItem").Value & "|@|" & vbCrLf  ' Mid(pRsDadosParaNFCe("Numero do Cheque").Value, 5, 2) & "|@|" & vbCrLf
                    
                    'CSTCSOSN
                    xStringNfce = xStringNfce & "863-" & Format(xOrdem, "000") & " = 0" & "|@|" & vbCrLf
                    

                    xStringNfce = xStringNfce & "899-" & Format(xOrdem, "000") & " = FIM" & 1 & "|@|" & vbCrLf
                    
                    pRsDadosParaNFCe.MoveNext
            Loop
            
            xStringNfce = xStringNfce & "999-999 = 0" & "|@|" & vbCrLf
            
            MontaTextoItensSolicitacaoNFCE = xStringNfce


End Function
Private Sub ImpTermicaAbreRelatorio()
    lNomeArquivo = BioCriaImprime
    'seleciona medidas para centímetros
    BioImprime "@@Printer.ScaleMode = 7"
    BioImprime "@@Printer.PaperSize = 1"
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@@Printer.FontName = Draft 10cpi"
    'teste para imprimir letra correta
    BioImprime "@@Printer.FontBold = False"
    BioImprime "@@ImprimeTexto " & Chr(34) & "  " & Chr(34) & ", 1, 2, 2, 1"
    BioImprime "@@Printer.FontName = Draft 5cpi"
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@@Printer.CurrentY = 0"
    BioImprime "@@Printer.Print " & "  "
    Printer.FontName = "Sans Serif 10cpi"
    Printer.FontName = "Lucida Console 7cpi"
    BioImprime "@@Printer.FontName = Lucida Console 7cpi"
    BioImprime "@@Printer.CurrentY = 0"
End Sub
Private Sub ImpTermicaImprimeDados(ByVal pLinhaDados As String, ByVal pNegrito As Boolean)
    Dim xNegrito As String
    
    If pNegrito = True Then
        xNegrito = "True"
    Else
        xNegrito = "False"
    End If
    BioImprime "@Printer.Print " & pLinhaDados
    BioImprime "@@Printer.FontBold = " & xNegrito
End Sub
Private Sub ImpTermicaCabecalhoPosto(ByVal pNumeroNFCe As Long)
    Dim xEmpresa As New cEmpresa
    Dim xLinhaDados As String
    Dim xDados As String
    Dim i As Integer
    
    If xEmpresa.LocalizarCodigo(g_empresa) = False Then
        Exit Sub
    End If
    Call ImpTermicaImprimeDados("------------------------------------------------", True)
    Call ImpTermicaImprimeDados(" ", True)
'    Call ImpTermicaImprimeDados(xEmpresa.Nome, True)
'    Call ImpTermicaImprimeDados("CNPJ: " & fMascaraCNPJ(xEmpresa.CGC) & "  IE: " & xEmpresa.InscricaoEstadual, True)
'    Call ImpTermicaImprimeDados(xEmpresa.Endereco, True)
'    Call ImpTermicaImprimeDados(xEmpresa.Cidade & "-" & xEmpresa.Estado, True)
'    Call ImpTermicaImprimeDados("CEP: " & fMascaraCEP(xEmpresa.CEP) & "  FONE: " & fMascaraTelefone(xEmpresa.Telefone), True)
    
    xLinhaDados = Space(48)
    i = Len(Trim(xEmpresa.Nome))
    Mid(xLinhaDados, 4 + ((40 - i) / 2), i) = Trim(xEmpresa.Nome)
    Call ImpTermicaImprimeDados(xLinhaDados, True)
    
    xLinhaDados = Space(48)
    xDados = "CNPJ: " & fMascaraCNPJ(xEmpresa.CGC) & "  IE: " & xEmpresa.InscricaoEstadual
    i = Len(Trim(xDados))
    Mid(xLinhaDados, 4 + ((40 - i) / 2), i) = Trim(xDados)
    Call ImpTermicaImprimeDados(xLinhaDados, True)
    
    xLinhaDados = Space(48)
    xDados = xEmpresa.Endereco
    i = Len(Trim(xDados))
    Mid(xLinhaDados, 4 + ((40 - i) / 2), i) = Trim(xDados)
    Call ImpTermicaImprimeDados(xLinhaDados, True)
    
    xLinhaDados = Space(48)
    xDados = xEmpresa.Cidade & "-" & xEmpresa.Estado
    i = Len(Trim(xDados))
    Mid(xLinhaDados, 4 + ((40 - i) / 2), i) = Trim(xDados)
    Call ImpTermicaImprimeDados(xLinhaDados, True)
    
    xLinhaDados = Space(48)
    xDados = "CEP: " & fMascaraCEP(xEmpresa.CEP) & "  FONE: " & fMascaraTelefone(xEmpresa.Telefone)
    i = Len(Trim(xDados))
    Mid(xLinhaDados, 4 + ((40 - i) / 2), i) = Trim(xDados)
    Call ImpTermicaImprimeDados(xLinhaDados, True)
    
    Call ImpTermicaImprimeDados(" ", True)
    Call ImpTermicaImprimeDados("------------------------------------------------", True)
    xLinhaDados = "RELATÓRIO GERENCIAL"
    If (pNumeroNFCe > 0) Then
        xLinhaDados = xLinhaDados & " REFERENTE A NFCe: " & Format(pNumeroNFCe, "###,##0")
    End If
    Call ImpTermicaImprimeDados(" ", True)
    Call ImpTermicaImprimeDados(xLinhaDados, True)
    Call ImpTermicaImprimeDados("DATA: " & Format(Date, "dd/MM/yyyy") & " AS " & Format(Now, "HH:mm:ss"), True)
    
    xLinhaDados = "Funcionario: 000                                "
    Mid(xLinhaDados, 14, 3) = Format(l_codigo_funcionario, "000")
    Mid(xLinhaDados, 18, Len(l_nome_funcionario)) = l_nome_funcionario
    Call ImpTermicaImprimeDados(xLinhaDados, True)
    
    Call ImpTermicaImprimeDados(" ", True)
    Call ImpTermicaImprimeDados("------------------------------------------------", True)
End Sub
Private Sub ImpTermicaDadosCliente()
    Dim xEmpresa As New cEmpresa
    Dim xLinhaDados As String
    Dim xStringParcial As String
    Dim i As Integer
    
    xLinhaDados = "CPF/CNPJ.:                                       "
    Mid(xLinhaDados, 11, 20) = txt_cpf.Text
    If Val(Cliente.InscricaoEstadual) > 0 Then
        Mid(xLinhaDados, 30, 19) = "IE:" & Mid(Cliente.InscricaoEstadual, 1, 16)
    End If
    Call ImpTermicaImprimeDados(xLinhaDados, True)
     
    xLinhaDados = "NOME..:                                         "
    Mid(xLinhaDados, 9, 40) = txt_nome_cliente.Text
    Call ImpTermicaImprimeDados(xLinhaDados, True)
        
    xLinhaDados = "END:                                            "
    Mid(xLinhaDados, 5, 44) = Cliente.Endereco
    Call ImpTermicaImprimeDados(xLinhaDados, True)
    
    If Len(Trim(Cliente.Bairro) & " - " & Trim(Cliente.Cidade) & " - " & Trim(Cliente.UF)) <= 48 Then
        xLinhaDados = "                                                "
        Mid(xLinhaDados, 5, 48) = Trim(Cliente.Bairro) & " - " & Trim(Cliente.Cidade) & " - " & Trim(Cliente.UF)
        Call ImpTermicaImprimeDados(xLinhaDados, True)
    Else
        xLinhaDados = "                                                "
        Mid(xLinhaDados, 5, 48) = Trim(Cliente.Bairro)
        Call ImpTermicaImprimeDados(xLinhaDados, True)
    
        xLinhaDados = "                                                "
        Mid(xLinhaDados, 5, 48) = Trim(Cliente.Cidade) & " - " & Trim(Cliente.UF)
        Call ImpTermicaImprimeDados(xLinhaDados, True)
    End If
    
    If lCodigoVeiculo > 0 Then
        xLinhaDados = "VEICULO:                                        "
        If Len(Trim(VeiculoCliente.Nome) & ", " & Trim(VeiculoCliente.Cor) & ", " & VeiculoCliente.ano & ", " & VeiculoCliente.PlacaLetra & "-" & VeiculoCliente.PlacaNumero) <= 40 Then
            Mid(xLinhaDados, 9, 40) = Trim(VeiculoCliente.Nome) & ", " & Trim(VeiculoCliente.Cor) & ", " & VeiculoCliente.ano & ", " & VeiculoCliente.PlacaLetra & "-" & VeiculoCliente.PlacaNumero
            Call ImpTermicaImprimeDados(xLinhaDados, True)
        Else
            Mid(xLinhaDados, 9, 40) = Trim(VeiculoCliente.Nome) & ", " & VeiculoCliente.ano
            Call ImpTermicaImprimeDados(xLinhaDados, True)
            
            xLinhaDados = "COR/PLC:                                        "
            Mid(xLinhaDados, 9, 40) = Trim(VeiculoCliente.Cor) & ", " & VeiculoCliente.PlacaLetra & "-" & VeiculoCliente.PlacaNumero
            Call ImpTermicaImprimeDados(xLinhaDados, True)
        End If
        If txt_kilometragem.Text <> "" Then
            xLinhaDados = "KILOMETRAGEM..:                                 "
            Mid(xLinhaDados, 17, 12) = txt_kilometragem.Text
            Call ImpTermicaImprimeDados(xLinhaDados, True)
        Else
            Call ImpTermicaImprimeDados("KILOMETRAGEM..: ______________                  ", True)
        End If
    Else
        Call ImpTermicaImprimeDados("Veiculo.: __________________________            ", True)
        Call ImpTermicaImprimeDados("                                                ", True)
        
        If txt_placa.Text <> "" Or txt_kilometragem.Text <> "" Then
            xLinhaDados = "PLACA.:             KILOMETRAGEM..:             "
            Mid(xLinhaDados, 9, 8) = txt_placa.Text
            Mid(xLinhaDados, 37, 12) = txt_kilometragem.Text
            Call ImpTermicaImprimeDados(xLinhaDados, True)
        Else
            Call ImpTermicaImprimeDados("Placa...: ______  ________   KM.:_____________  ", True)
        End If
    End If
    
    
    
    
    Call ImpTermicaImprimeProdutos(lNumeroCupom, lData)
    Call ImpTermicaImprimeDados("------------------------------------------------", True)
    Call ImpTermicaImprimeDados("                                                ", True)
    Call ImpTermicaImprimeDados("     Recebi(emos) a(s) mercadoria(s) da NFCe de ", True)
    Call ImpTermicaImprimeDados("N. " & Format(lNumeroCupom, "###,##0") & " e Pagarei(emos) a Importância acima. ", True)
    Call ImpTermicaImprimeDados("                                                ", True)
    Call ImpTermicaImprimeDados("   X________________________________________    ", True)
    
    xLinhaDados = Space(48)
    i = Len(Trim(Cliente.RazaoSocial))
    Mid(xLinhaDados, 4 + ((40 - i) / 2), i) = Trim(Cliente.RazaoSocial)
    Call ImpTermicaImprimeDados(xLinhaDados, True)
    Call ImpTermicaImprimeDados(" ", True)
    Call ImpTermicaImprimeDados(" ", True)
    Call ImpTermicaImprimeDados(" ", True)
End Sub

Private Function MontaTextoProdutosNotaAbastecimento(ByVal pNumeroCupom As String, ByVal pData As Date) As String



Dim xConteudo As String
Dim i As Integer
Dim i2 As Integer
Dim xOrdem As Integer
Dim xStringAuxiliar As String
Dim xStringAuxiliar2 As String
Dim xMovCupomFiscal As New cMovimentoCupomFiscal
Dim xTotalProdutos As Currency

xOrdem = 0
xConteudo = ""
xStringAuxiliar = ""
i2 = 0

xTotalProdutos = 0
Const TEXTO_TOTAL As String = "T O T A L"

MontaTextoProdutosNotaAbastecimento = ""

    Do Until xMovCupomFiscal.LocalizarNumeroProximaOrdem(g_empresa, lNumeroPDV, pNumeroCupom, pData, xOrdem) = False
        i = i + 1
        If i = 1 Then
            
            xConteudo = xConteudo + "------------------------------------------------" + Chr(13) + Chr(10)
            xConteudo = xConteudo + "ITEM   CÓDIGO             DESCRIÇÃO             " + Chr(13) + Chr(10)
            xConteudo = xConteudo + "      QTDxUNITÁRIO       ST          VALOR( R$) " + Chr(13) + Chr(10)
            xConteudo = xConteudo + "------------------------------------------------" + Chr(13) + Chr(10)
        
        End If


        xStringAuxiliar = Space(48)
        Mid(xStringAuxiliar, 1, 3) = Format(xMovCupomFiscal.Ordem, "000")
        Mid(xStringAuxiliar, 5, 4) = Format(xMovCupomFiscal.CodigoProduto, "###0")

        If Produto.LocalizarCodigo(xMovCupomFiscal.CodigoProduto) Then
            Mid(xStringAuxiliar, 10, 40) = Produto.Nome
        End If
        xConteudo = xConteudo + xStringAuxiliar + Chr(13) + Chr(10)

        xStringAuxiliar = Space(48)

        i2 = Len(Format(xMovCupomFiscal.Quantidade, "##########0.000"))
        Dim xPosicaoInicialValor As Integer
        xPosicaoInicialValor = 1 + 7 - i2
        'Mid(xStringAuxiliar, 1 + 7 - i2, i2) = Format(xMovCupomFiscal.Quantidade, "##########0.000")
        Mid(xStringAuxiliar, IIf(xPosicaoInicialValor < 1, 1, xPosicaoInicialValor), i2) = Format(xMovCupomFiscal.Quantidade, "##########0.000")


        Mid(xStringAuxiliar, 8, 3) = Mid(Produto.Unidade, 1, 2) + "x"
        xStringAuxiliar2 = Format(xMovCupomFiscal.ValorUnitario, "00000000000.000")
        If Mid(xStringAuxiliar2, 15, 1) = 0 Then
            Mid(xStringAuxiliar, 11, 15) = Format(xMovCupomFiscal.ValorUnitario, "###########0.00")
        Else
            Mid(xStringAuxiliar, 11, 15) = Format(xMovCupomFiscal.ValorUnitario, "##########0.000")
        End If
        If Aliquota.LocalizarCodigo(lSerieECF, xMovCupomFiscal.CodigoAliquota) Then
            Mid(xStringAuxiliar, 26, 2) = Aliquota.CodigoFiscal
        End If
        i2 = Len(Format(xMovCupomFiscal.ValorTotal, "###########0.00"))
        Mid(xStringAuxiliar, 33 + 15 - i2, i2) = Format(xMovCupomFiscal.ValorTotal, "###########0.00")

        xConteudo = xConteudo + xStringAuxiliar + Chr(13) + Chr(10)
        xOrdem = xOrdem + 1
        
        xTotalProdutos = xTotalProdutos + xMovCupomFiscal.ValorTotal
    Loop
    
    xStringAuxiliar = Space(48)
    If xOrdem > 0 Then
        xConteudo = xConteudo + "------------------------------------------------" + Chr(13) + Chr(10)
        Mid(xStringAuxiliar, 1, Len(TEXTO_TOTAL)) = TEXTO_TOTAL
        i2 = Len(Format(xTotalProdutos, "###########0.00"))
        Mid(xStringAuxiliar, 33 + 15 - i2, i2) = Format(xTotalProdutos, "###########0.00")
        xConteudo = xConteudo + xStringAuxiliar + Chr(13) + Chr(10)
    End If


MontaTextoProdutosNotaAbastecimento = xConteudo


End Function

Private Sub ImpTermicaImprimeProdutos(ByVal pNumeroCupom As String, ByVal pData As Date)
    Dim xLinhasCupom As Variant
    Dim xLinhaDados As String
    Dim i As Integer
    Dim i2 As Integer
        
        
'    xLinhasCupom = Split(txt_cupom_fiscal.Text, vbCrLf)
'    For i = 0 To UBound(xLinhasCupom) - 1
'        If i >= 4 Then
'            xLinhaDados = Space(48)
'            i2 = Len(xLinhasCupom(i))
'            Mid(xLinhaDados, 1, i2) = xLinhasCupom(i)
'            If Mid(xLinhaDados, 1, 10) = "T O T A L " Then
'                Call ImpTermicaImprimeDados(xLinhaDados, True)
'                Exit For
'            End If
'           Call ImpTermicaImprimeDados(xLinhaDados, True)
'        End If
'    Next

    xLinhasCupom = Split(MontaTextoProdutosNotaAbastecimento(pNumeroCupom, pData), vbCrLf)

    For i = 0 To UBound(xLinhasCupom) - 1
           xLinhaDados = Space(48)
           i2 = Len(xLinhasCupom(i))
           Mid(xLinhaDados, 1, i2) = xLinhasCupom(i)
           Call ImpTermicaImprimeDados(xLinhaDados, True)
    Next
End Sub
Private Sub ImpTermicaFechaRelatorio(ByVal pNomeRelatorio)
    BioImprime "@Printer.Print  "
    BioImprime "@Printer.Print  "
    BioImprime "@@Printer.EndDoc"
    BioFechaImprime
    g_string = lLocalImpressao & lNomeArquivo & "|@|" & pNomeRelatorio & "|@|"
    frm_preview.Show 1
End Sub
Private Function RetornaEncerranteInicial(pEncerranteFinal As Currency, pQuantidadeSaida As Currency)
    If pEncerranteFinal <= 0 Then
        RetornaEncerranteInicial = 0
        Exit Function
    End If
    RetornaEncerranteInicial = pEncerranteFinal - pQuantidadeSaida
End Function
'
'
'
Private Function RetornaValorImpostoProdutoNFCE(ByVal pValorBaseCalculo As Currency, ByVal pAliquotaImposto As Currency) As Currency

    If Val(pAliquotaImposto) = 0 Then
        RetornaValorImpostoProdutoNFCE = pAliquotaImposto
        Exit Function
    End If
    
    Dim xAliquota As Currency
    
    xAliquota = pAliquotaImposto / 100
    
    RetornaValorImpostoProdutoNFCE = pValorBaseCalculo * xAliquota

End Function

Private Function MontaTextoCabecalhoSolicitacaoNFCE(ByVal pRsDadosParaNFCe As ADODB.Recordset, ByVal pTipoServico As String) As String

    MontaTextoCabecalhoSolicitacaoNFCE = Empty
    
    Dim xEmpresa As New cEmpresa
    Dim xStringNfce As String
    Dim xTipoServico As String
    Dim xCNPJEmpresa As String
    Dim xCodigoCidadeIbgeEmitente As String
    Dim xCodigoCidadeIbgeDestinatario As String
    
    Dim xSuframaCliente As String
    Dim xInscricaoEstadualCliente As String
    Dim xEmailCliente As String
    Dim xCidadeCliente As String
    Dim xEnderecoCliente As String
    Dim xBairroCliente As String
    Dim xUFCliente As String
    Dim xCEPCliente As String
    Dim xTelefoneCliente As String
    Dim xNomeCliente As String
    Dim xCPFCNPJCliente As String
    Dim xIndicadorIEDestinatario As Integer
    Const IE_ISENTO As String = "ISENT*"
    
    Dim xIdentificadorLocalOperacao As Integer
    
    xTipoServico = pTipoServico
    xCNPJEmpresa = Empty
    
    xIdentificadorLocalOperacao = 1 'operacao interna
    
    
    
    If xEmpresa.LocalizarCodigo(lEmpresa) Then xCNPJEmpresa = xEmpresa.CGC
    If CidadeIBGE.LocalizarNome(UCase(xEmpresa.Estado), UCase(xEmpresa.Cidade)) Then
        xCodigoCidadeIbgeEmitente = CidadeIBGE.Codigo
        xCodigoCidadeIbgeDestinatario = CidadeIBGE.Codigo
    End If
    
    xIndicadorIEDestinatario = 9
    If Val(pRsDadosParaNFCe("IdClienteFornecedor_MovDECabecalho").Value) > 0 Then
        If Cliente.LocalizarCodigo(pRsDadosParaNFCe("IdClienteFornecedor_MovDECabecalho").Value) Then
            If CidadeIBGE.LocalizarNome(UCase(Cliente.UF), UCase(Cliente.Cidade)) Then
                xCodigoCidadeIbgeDestinatario = CidadeIBGE.Codigo
            End If
            If Cliente.ImprimeDadosECF = True Then
                xNomeCliente = Cliente.RazaoSocial
                xEmailCliente = Cliente.Email
                xCidadeCliente = Cliente.Cidade
                xEnderecoCliente = Cliente.Endereco
                xBairroCliente = Cliente.Bairro
                xSuframaCliente = Empty
                xUFCliente = Cliente.UF
                xCEPCliente = Cliente.CEP
                xTelefoneCliente = Cliente.Telefone
                xInscricaoEstadualCliente = Cliente.InscricaoEstadual
                If Cliente.CGC <> "" Then
                    xCPFCNPJCliente = Cliente.CGC
                Else
                    xCPFCNPJCliente = Cliente.CPF
                End If
                
'                If Trim(UCase(Cliente.UF)) <> Trim(UCase(xEmpresa.Estado)) Then
'                    xIdentificadorLocalOperacao = 2
'                End If
                
                'Foi levantado que para emitir NFCe, sempre usa codigo 9
                'If UCase(Trim(xInscricaoEstadualCliente)) Like IE_ISENTO Then
                '    xIndicadorIEDestinatario = 2
                'ElseIf Len(Trim(xInscricaoEstadualCliente)) > 0 Then
                '    If IsNumeric(Mid(Trim(xInscricaoEstadualCliente), 1, 1)) Then
                '        If Cliente.Contribuinte = True Then
                '            xIndicadorIEDestinatario = 1
                '        End If
                '    End If
                'End If
            End If
        End If
    Else
        If Len(txt_cpf.Text) = 14 Then
            xCPFCNPJCliente = fDesmascaraCPF(txt_cpf.Text)
        ElseIf Len(txt_cpf.Text) = 18 Then
            xCPFCNPJCliente = fDesmascaraCNPJ(txt_cpf.Text)
        End If
        If Len(txt_nome_cliente.Text) > 0 Then
            xNomeCliente = txt_nome_cliente.Text
        End If
    End If
    

    xStringNfce = xStringNfce & "000-000 = " & xTipoServico & "|@|" & vbCrLf
    
    'NÃO NECESSÁRIO PARA NFCE MANTIDO APENAS PARA CONSERVAR PADRÃO
    xStringNfce = xStringNfce & "001-000 = " & Format(0, "0000000000") & "|@|" & vbCrLf  'OBTER DADOS REAIS - MÉTODO PARA GERA CONTROLE DE SOLICITAÇÃO
    xStringNfce = xStringNfce & "002-000 = " & gVersaoSGP & "|@|" & vbCrLf
    xStringNfce = xStringNfce & "040-000 = " & g_empresa & "|@|" & vbCrLf 'OBTER DADOS REAIS
    xStringNfce = xStringNfce & "041-000 = " & "TEIXEIRA E PINHEIRO LTDA" & "|@|" & vbCrLf 'OBTER DADOS REAIS
    xStringNfce = xStringNfce & "042-000 = " & "3" & "|@|" & vbCrLf 'CRT 1-SIMPLES NACIONAL, 3-REGIME NORMAL
    xStringNfce = xStringNfce & "045-000 = " & xIdentificadorLocalOperacao & "|@|" & vbCrLf 'Local de Destino da operação: 1-Interna, 2-Interestadual, 3-Exterior
    xStringNfce = xStringNfce & "046-000 = " & "1" & "|@|" & vbCrLf 'FINALIDADE DA NFE
    
    
    'NUMERAÇÃO ESTÁ FORA DA SEQUENCIA POIS INICIALMENTE FOI UTILIZADA A ESTRUTURA EXISTENTE DA SOLICITAÇÃO DE NFE
    'ALGUMAS NUMERAÇÕES FORAM REMOVIDAS POR NÃO TEREM UTILIDADE NA NFCE
    'QUALQUER ALTERAÇÃO NESTAS NUMERAÇÕES IMPACTAM O FUNCIONAMENTO DA EMISSÃO NFCE PELA DLL
    xStringNfce = xStringNfce & "100-000 = INICIO" & "|@|" & vbCrLf 'VALOR FIXO
    
    xStringNfce = xStringNfce & "111-000 = " & lNumeroCupom & "|@|" & vbCrLf

    xStringNfce = xStringNfce & "112-000 = " & lSerieNFCe & "|@|" & vbCrLf

'---- DADOS DO PAGAMENTO -----
    'Aqui somente Cartão Débito e Cartão Crédito
    If lNFCe_tPag <> "" Then
        xStringNfce = xStringNfce & "113-001 = " & lNFCe_tPag & "|@|" & vbCrLf 'Forma de Pagamento
        
        xStringNfce = xStringNfce & "114-001 = " & lNFCe_TpIntegra & "|@|" & vbCrLf 'Tipo de integração do Pagamento 1=Pagamento integrado com o sistema 2=equipamento POS
        
        xStringNfce = xStringNfce & "115-001 = " & lNFCe_CNPJCartao & "|@|" & vbCrLf 'CNPJ OPERADORA DO CARTÃO
    
        xStringNfce = xStringNfce & "116-001 = " & FormatNumber(lNFCe_vPag, 2) & "|@|" & vbCrLf
    
        xStringNfce = xStringNfce & "117-001 = " & lNFCe_tBand & "|@|" & vbCrLf 'BANDEIRA DO CARTÃO
        
        xStringNfce = xStringNfce & "118-001 = " & lNFCe_cAut & "|@|" & vbCrLf 'NUMERO AUTORIZAÇÃO DO CARTÃO (OBRIGATÓRIO SE INFORMAR O CNPJ DA OPERADORA)
        
        xStringNfce = xStringNfce & "119-001 = " & "FIM" & "|@|" & vbCrLf
    Else
        Dim xFormaPagamento As String
        
        '1-Dinheiro
        If pRsDadosParaNFCe("FormaPagamento_MovDECabecalho").Value = 1 Then
            xFormaPagamento = "01" '01-Diheiro
        '2-Cheque à Vista, 3-Cheque Pré-Datado
        ElseIf pRsDadosParaNFCe("FormaPagamento_MovDECabecalho").Value = 2 Or pRsDadosParaNFCe("FormaPagamento_MovDECabecalho").Value = 3 Then
            xFormaPagamento = "02" '02-Cheque
        '4-Nota Vinculada, 17-Cerrado Tef
        ElseIf pRsDadosParaNFCe("FormaPagamento_MovDECabecalho").Value = 5 Or pRsDadosParaNFCe("FormaPagamento_MovDECabecalho").Value = 17 Then
            xFormaPagamento = "05" '05-Crédito Loja
        Else
            xFormaPagamento = "99" '99-Outros
        End If
        
        
        xStringNfce = xStringNfce & "113-001 = " & xFormaPagamento & "|@|" & vbCrLf 'Forma de Pagamento
        
        xStringNfce = xStringNfce & "114-001 = " & "|@|" & vbCrLf 'Tipo de integração do Pagamento 1=Pagamento integrado com o sistema 2=equipamento POS
        
        xStringNfce = xStringNfce & "115-001 = " & "|@|" & vbCrLf 'CNPJ OPERADORA DO CARTÃO
    
        xStringNfce = xStringNfce & "116-001 = " & FormatNumber(pRsDadosParaNFCe("ValorTotal_MovDECabecalho").Value, 2) & "|@|" & vbCrLf
    
        xStringNfce = xStringNfce & "117-001 = " & "|@|" & vbCrLf 'BANDEIRA DO CARTÃO
        
        xStringNfce = xStringNfce & "118-001 = " & "|@|" & vbCrLf 'NUMERO AUTORIZAÇÃO DO CARTÃO (OBRIGATÓRIO SE INFORMAR O CNPJ DA OPERADORA)
        
        xStringNfce = xStringNfce & "119-001 = " & "FIM" & "|@|" & vbCrLf
    End If
    
    xStringNfce = xStringNfce & "134-000 = " & "|@|" & vbCrLf 'pRsDadosParaNFCe("Codigo da Ecf").Value & "|@|" & vbCrLf 'NUMERO DA ECF
     
    'LOGICA FOI IMPLEMENTADA NO MÉTODO MontaTextoInformacoesComplementaresNFCe
'    If lDescontoPostoAki = True And lValorDescontoConcedido > 0 Then
'        xStringNfce = xStringNfce & "135-000 = " & MontaTextoInformacoesComplementaresNFCe & "Desconto pelo Aplicativo PostoAki" & "\n" & "|@|" & vbCrLf 'InfCpl - Informações complementares
'    Else
'        xStringNfce = xStringNfce & "135-000 = " & MontaTextoInformacoesComplementaresNFCe & "|@|" & vbCrLf 'InfCpl - Informações complementares
'    End If

    xStringNfce = xStringNfce & "135-000 = " & MontaTextoInformacoesComplementaresNFCe & "|@|" & vbCrLf 'InfCpl - Informações complementares


    

'---- 'OBTER DADOS REAIS DO 210 AO 223 dados do cliente ---

    xStringNfce = xStringNfce & "210-000 = " & xNomeCliente & "|@|" & vbCrLf 'pRsDadosParaNFCe("Nome").Value & "|@|" & vbCrLf  '& RetiraAcentos(Cliente.RazaoSocial) & "|@|" & vbCrLf
            
    xStringNfce = xStringNfce & "211-000 = " & xEnderecoCliente & "|@|" & vbCrLf
            
    xStringNfce = xStringNfce & "212-000 = " & xBairroCliente & "|@|" & vbCrLf
            
    xStringNfce = xStringNfce & "213-000 = " & xCidadeCliente & "|@|" & vbCrLf 'RetiraAcentos(Cliente.Cidade) & "|@|" & vbCrLf
            
    xStringNfce = xStringNfce & "214-000 = " & xUFCliente & "|@|" & vbCrLf 'RetiraAcentos(Cliente.UF.ToUpper) & "|@|" & vbCrLf
    
    xStringNfce = xStringNfce & "215-000 = " & xCEPCliente & "|@|" & vbCrLf 'Cliente.CEP & "|@|" & vbCrLf
    
    
    xStringNfce = xStringNfce & "216-000 = 1058" & "|@|" & vbCrLf 'Codigo País
    
    xStringNfce = xStringNfce & "217-000 = BRASIL" & "|@|" & vbCrLf 'RetiraAcentos("BRASIL") & "|@|" & vbCrLf
    
    xStringNfce = xStringNfce & "218-000 = " & xTelefoneCliente & "|@|" & vbCrLf '& Cliente.Telefone & "|@|" & vbCrLf
    
    xStringNfce = xStringNfce & "219-000 = " & xCPFCNPJCliente & "|@|" & vbCrLf 'pRsDadosParaNFCe("CPF CNPJ").Value & "|@|" & vbCrLf 'Cliente.CGC & "|@|" & vbCrLf 'cnpj

    xStringNfce = xStringNfce & "220-000 = " & xInscricaoEstadualCliente & "|@|" & vbCrLf  'inscrição estadual
    
    xStringNfce = xStringNfce & "221-000 = " & "" & "|@|" & vbCrLf 'Codigo Suframa Cliente

    'xStringNfce = xStringNfce & "222-000 = 5208707" & "|@|" & vbCrLf 'CidadeIBGE.Codigo & "|@|" & vbCrLf 'CodMunCli GOIANIA
    xStringNfce = xStringNfce & "222-000 = " & xCodigoCidadeIbgeDestinatario & "|@|" & vbCrLf 'CidadeIBGE.Codigo & "|@|" & vbCrLf 'CodMunCli MORRINHOS
            
    xStringNfce = xStringNfce & "223-000 = " & xEmailCliente & "|@|" & vbCrLf
    
    '1-Contribuinte , 2-ISENTO, 9-Não Contribuinte
    xStringNfce = xStringNfce & "224-000 = " & xIndicadorIEDestinatario & "|@|" & vbCrLf

'---- DADOS DA EMPRESA EMITENTE ----

     xStringNfce = xStringNfce & "310-000 = " & xEmpresa.Nome & "|@|" & vbCrLf  'RetiraAcentos(Empresa.Nome) & vbCrLf 'RazaoSocialEmp

     xStringNfce = xStringNfce & "311-000 = " & xEmpresa.Nome & "|@|" & vbCrLf  'RetiraAcentos(Empresa.Nome) & vbCrLf 'NomeFantasiaEmp
                
     xStringNfce = xStringNfce & "312-000 = " & xEmpresa.Endereco & "|@|" & vbCrLf  'Logradouro
                
     xStringNfce = xStringNfce & "313-000 = " & "0" & "|@|" & vbCrLf  'NumLgrEmp
     
     xStringNfce = xStringNfce & "314-000 = " & "" & "|@|" & vbCrLf  'CplEmp
     
     xStringNfce = xStringNfce & "315-000 = " & xEmpresa.Bairro & "|@|" & vbCrLf  '& RetiraAcentos(Empresa.Bairro) & vbCrLf 'BairroEmp
     
     xStringNfce = xStringNfce & "316-000 = " & xEmpresa.CEP & "|@|" & vbCrLf  'Empresa.CEP & vbCrLf 'CepEmp
     
     xStringNfce = xStringNfce & "317-000 = " & xEmpresa.Telefone & "|@|" & vbCrLf  '& Empresa.Telefone & vbCrLf 'TelefoneEmp
     
     xStringNfce = xStringNfce & "318-000 = " & xEmpresa.InscricaoEstadual & "|@|" & vbCrLf  'Empresa.InscricaoEstadual IEEmp
                
     xStringNfce = xStringNfce & "319-000 = " & xEmpresa.Cidade & "|@|" & vbCrLf  '& RetiraAcentos(Empresa.Cidade) & vbCrLf 'NomeMunEmp
                
     xStringNfce = xStringNfce & "320-000 = " & xCodigoCidadeIbgeEmitente & "|@|" & vbCrLf  '& xCodigoCidadeEmitente & vbCrLf 'CodMunEmp
                
     xStringNfce = xStringNfce & "321-000 = " & "" & "|@|" & vbCrLf  'CodMunIdentificacao
     
     xStringNfce = xStringNfce & "322-000 = " & "" & "|@|" & vbCrLf  'Empresa.EmailContador.ToLower) 'Email Contador(a)
     
     xStringNfce = xStringNfce & "323-000 = " & xCNPJEmpresa & "|@|" & vbCrLf ' 05577906000197 & Empresa.CGC & "|@|" & vbCrLf 'CNPJ
     
     xStringNfce = xStringNfce & "324-000 = " & xEmpresa.Estado & "|@|" & vbCrLf 'UF EMPRESA - ADICIONADO PARA NFCE
     
     xStringNfce = xStringNfce & "325-000 = 1" & "|@|" & vbCrLf 'TIPO NF (0-ENTRADA 1-SAIDA)
     
     xStringNfce = xStringNfce & "326-000 = VENDA" & "|@|" & vbCrLf 'NATUREZA OPERACAO
     
     xStringNfce = xStringNfce & "350-000 = " & "" & "|@|" & vbCrLf 'CnpjOuCpfTranspor
     xStringNfce = xStringNfce & "351-000 = " & "" & "|@|" & vbCrLf 'RazaoSocialTranspor
     xStringNfce = xStringNfce & "352-000 = " & "" & "|@|" & vbCrLf 'IETranspor
     xStringNfce = xStringNfce & "353-000 = " & "" & "|@|" & vbCrLf 'EndTranspor
     xStringNfce = xStringNfce & "354-000 = " & "" & "|@|" & vbCrLf 'NomeMunTranspor
     xStringNfce = xStringNfce & "355-000 = " & "" & "|@|" & vbCrLf 'UFTranspor
     xStringNfce = xStringNfce & "356-000 = " & "" & "|@|" & vbCrLf 'PlacaTranspor
     xStringNfce = xStringNfce & "357-000 = " & "" & "|@|" & vbCrLf 'UfPlacaTranspor
     xStringNfce = xStringNfce & "358-000 = 0" & "|@|" & vbCrLf '& FormatNumber(lTotalQtd, 0) & "|@|" & vbCrLf 'QntTranspor
     xStringNfce = xStringNfce & "359-000 = " & "|@|" & vbCrLf 'EspecieTranspor
     xStringNfce = xStringNfce & "360-000 = 0" & "|@|" & vbCrLf 'PesoLiqTranspor
     xStringNfce = xStringNfce & "361-000 = 0" & "|@|" & vbCrLf 'PesoBrutoTranspor
     xStringNfce = xStringNfce & "362-000 = 9" & "|@|" & vbCrLf 'TipoFrete
     xStringNfce = xStringNfce & "363-000 = " & "" & "|@|" & vbCrLf 'CodMunTranspor
     xStringNfce = xStringNfce & "364-000 = 0" & "|@|" & vbCrLf 'BCICMSTransp
     xStringNfce = xStringNfce & "365-000 = 0" & "|@|" & vbCrLf 'AliquotaIcmsTransp
     xStringNfce = xStringNfce & "366-000 = 0" & "|@|" & vbCrLf 'ValorICMSTransp
     xStringNfce = xStringNfce & "367-000 = " & "|@|" & vbCrLf 'CFOPTransp
     xStringNfce = xStringNfce & "368-000 = 0" & "|@|" & vbCrLf 'ValorServicoMCC

     
     
    
    MontaTextoCabecalhoSolicitacaoNFCE = xStringNfce

End Function

Private Function AtualizaTabelaSolicitacaoNFCe(ByVal pTipoOperacao As String, ByVal pChaveAcessoNFe As String, ByVal pTexto As String, ByVal pNumeroDaNota As Long, ByVal pNumeroLote As String) As Boolean
    AtualizaTabelaSolicitacaoNFCe = False
    
    Call CriaLogECF(Time & " [AtualizaTabelaSolicitacaoNFCe] pTipoOperacao=" & pTipoOperacao & " - pNumeroDaNota=" & pNumeroDaNota & " - pChaveAcessoNFe=" & pChaveAcessoNFe)
    Set MovSolicitacaoFuncaoNFe = New cMovSolicitacaoFuncaoNFe
    
    MovSolicitacaoFuncaoNFe.NSU_MovSolicitacaoFuncaoNFe = 0 'No momento da inserção é feita a busca para obter o proximo registro
    MovSolicitacaoFuncaoNFe.NumeroControleSolicitacao_MovSolicitacaoFuncaoNFe = 0
    MovSolicitacaoFuncaoNFe.DataSolicitacao_MovSolicitacaoFuncaoNFe = CDate(Format(Now, "dd-MM-yyyy"))
    'MovSolicitacaoFuncaoNFe.HoraSolicitacao = CDate(Format(Now, "HH:mm:ss"))
    MovSolicitacaoFuncaoNFe.TipoOperacao_MovSolicitacaoFuncaoNFe = pTipoOperacao
    MovSolicitacaoFuncaoNFe.CodigoEstabelecimento_MovSolicitacaoFuncaoNFe = g_empresa
    MovSolicitacaoFuncaoNFe.SerieNFe_MovSolicitacaoFuncaoNFe = lSerieNFCe 'Verficar se pode ser utilizado este número
    MovSolicitacaoFuncaoNFe.NumeroNFe_MovSolicitacaoFuncaoNFe = pNumeroDaNota
    MovSolicitacaoFuncaoNFe.ChaveAcessoNFe_MovSolicitacaoFuncaoNFe = pChaveAcessoNFe
    If pTipoOperacao = "STATUS SERVICO" Or pTipoOperacao = "ATV" Or pTipoOperacao = "IMPRESSAO" Then
        MovSolicitacaoFuncaoNFe.SerieNFe_MovSolicitacaoFuncaoNFe = ""
        MovSolicitacaoFuncaoNFe.NumeroNFe_MovSolicitacaoFuncaoNFe = "0"
    End If
    MovSolicitacaoFuncaoNFe.IPComputadorAC_MovSolicitacaoFuncaoNFe = GetIPAddress()
    MovSolicitacaoFuncaoNFe.IPInternetAC_MovSolicitacaoFuncaoNFe = "200??.??.??.??"
    MovSolicitacaoFuncaoNFe.SegurancaEstabelecimento_MovSolicitacaoFuncaoNFe = "1234"
    MovSolicitacaoFuncaoNFe.CodigoUsuario_MovSolicitacaoFuncaoNFe = g_usuario
    MovSolicitacaoFuncaoNFe.VersaoAC_MovSolicitacaoFuncaoNFe = gVersaoSGP
    MovSolicitacaoFuncaoNFe.VersaoHost_MovSolicitacaoFuncaoNFe = "??"
    MovSolicitacaoFuncaoNFe.Texto_MovSolicitacaoFuncaoNFe = pTexto
    MovSolicitacaoFuncaoNFe.HoraAnalise_MovSolicitacaoFuncaoNFe = CDate("00:00:00")
    MovSolicitacaoFuncaoNFe.HoraAprovacao_MovSolicitacaoFuncaoNFe = CDate("00:00:00")
    MovSolicitacaoFuncaoNFe.HoraCancelamentoHost_MovSolicitacaoFuncaoNFe = CDate("00:00:00")
    MovSolicitacaoFuncaoNFe.HoraConfirmacaoAC_MovSolicitacaoFuncaoNFe = CDate("00:00:00")
    MovSolicitacaoFuncaoNFe.HoraCancelamentoAC_MovSolicitacaoFuncaoNFe = CDate("00:00:00")
    MovSolicitacaoFuncaoNFe.Mensagem_MovSolicitacaoFuncaoNFe = ""
    MovSolicitacaoFuncaoNFe.CodigoRetorno_MovSolicitacaoFuncaoNFe = 0
    MovSolicitacaoFuncaoNFe.NumeroLote_MovSolicitacaoFuncaoNFe = pNumeroLote
    
    AtualizaTabelaSolicitacaoNFCe = LoopIncluiRegistroSolicitacaoNFe()

End Function
Private Function GravaSolicitacaoProcessamentoNFCe(ByVal pNumeroNFCe As String, ByVal pStringChamadaProcessamento As String) As Boolean
    GravaSolicitacaoProcessamentoNFCe = False
    
    On Error GoTo TrataError
    
    Dim xMovSolicitacaoFuncaoNFe As New cMovSolicitacaoFuncaoNFe
    
    xMovSolicitacaoFuncaoNFe.NSU_MovSolicitacaoFuncaoNFe = 0 'No momento da inserção é feita a busca para obter o proximo registro
    xMovSolicitacaoFuncaoNFe.NumeroControleSolicitacao_MovSolicitacaoFuncaoNFe = 0
    xMovSolicitacaoFuncaoNFe.DataSolicitacao_MovSolicitacaoFuncaoNFe = CDate(Format(Now, "dd-MM-yyyy"))
    xMovSolicitacaoFuncaoNFe.TipoOperacao_MovSolicitacaoFuncaoNFe = "PROCESSA_NFCE"
    xMovSolicitacaoFuncaoNFe.CodigoEstabelecimento_MovSolicitacaoFuncaoNFe = g_empresa
    xMovSolicitacaoFuncaoNFe.SerieNFe_MovSolicitacaoFuncaoNFe = lSerieNFCe 'Verficar se pode ser utilizado este número
    xMovSolicitacaoFuncaoNFe.NumeroNFe_MovSolicitacaoFuncaoNFe = pNumeroNFCe
    xMovSolicitacaoFuncaoNFe.ChaveAcessoNFe_MovSolicitacaoFuncaoNFe = ""
    xMovSolicitacaoFuncaoNFe.IPComputadorAC_MovSolicitacaoFuncaoNFe = GetIPAddress()
    xMovSolicitacaoFuncaoNFe.IPInternetAC_MovSolicitacaoFuncaoNFe = "200??.??.??.??"
    xMovSolicitacaoFuncaoNFe.SegurancaEstabelecimento_MovSolicitacaoFuncaoNFe = "1234"
    xMovSolicitacaoFuncaoNFe.CodigoUsuario_MovSolicitacaoFuncaoNFe = g_usuario
    xMovSolicitacaoFuncaoNFe.VersaoAC_MovSolicitacaoFuncaoNFe = gVersaoSGP
    xMovSolicitacaoFuncaoNFe.VersaoHost_MovSolicitacaoFuncaoNFe = "??"
    xMovSolicitacaoFuncaoNFe.Texto_MovSolicitacaoFuncaoNFe = pStringChamadaProcessamento
    xMovSolicitacaoFuncaoNFe.HoraAnalise_MovSolicitacaoFuncaoNFe = CDate("00:00:00")
    xMovSolicitacaoFuncaoNFe.HoraAprovacao_MovSolicitacaoFuncaoNFe = CDate("00:00:00")
    xMovSolicitacaoFuncaoNFe.HoraCancelamentoHost_MovSolicitacaoFuncaoNFe = CDate("00:00:00")
    xMovSolicitacaoFuncaoNFe.HoraConfirmacaoAC_MovSolicitacaoFuncaoNFe = CDate("00:00:00")
    xMovSolicitacaoFuncaoNFe.HoraCancelamentoAC_MovSolicitacaoFuncaoNFe = CDate("00:00:00")
    xMovSolicitacaoFuncaoNFe.Mensagem_MovSolicitacaoFuncaoNFe = ""
    xMovSolicitacaoFuncaoNFe.CodigoRetorno_MovSolicitacaoFuncaoNFe = 0
    xMovSolicitacaoFuncaoNFe.NumeroLote_MovSolicitacaoFuncaoNFe = 0
    
    GravaSolicitacaoProcessamentoNFCe = xMovSolicitacaoFuncaoNFe.Incluir
    
    Exit Function
TrataError:
    Call CriaLogSGP("[GravaSolicitacaoProcessamentoNFCe]", "Erro ao tentar gravar solicitação do processamento da NFCe - " & Err.Description, "pStringChamadaProcessamento=" & pStringChamadaProcessamento)
    MsgBox "Não foi possível incluir registro de Solicitação do processamento para NFC-e!", vbCritical, "Erro de Integridade"
End Function
'Private Function GravaSolicitacaoComunicacaoPetromovelAuto() As Long
'
'    GravaSolicitacaoComunicacaoPetromovelAuto = 0
'
'    On Error GoTo TrataError
'
'    Dim xMovSolicitacaoFuncaoNFe As New cMovSolicitacaoFuncaoNFe
'
'    xMovSolicitacaoFuncaoNFe.NSU_MovSolicitacaoFuncaoNFe = 0 'No momento da inserção é feita a busca para obter o proximo registro
'    xMovSolicitacaoFuncaoNFe.NumeroControleSolicitacao_MovSolicitacaoFuncaoNFe = 0
'    xMovSolicitacaoFuncaoNFe.DataSolicitacao_MovSolicitacaoFuncaoNFe = CDate(Format(Now, "dd-MM-yyyy"))
'    xMovSolicitacaoFuncaoNFe.TipoOperacao_MovSolicitacaoFuncaoNFe = "PETROMOVEL_ATIVO"
'    xMovSolicitacaoFuncaoNFe.CodigoEstabelecimento_MovSolicitacaoFuncaoNFe = g_empresa
'    xMovSolicitacaoFuncaoNFe.SerieNFe_MovSolicitacaoFuncaoNFe = "0" 'Verficar se pode ser utilizado este número
'    xMovSolicitacaoFuncaoNFe.NumeroNFe_MovSolicitacaoFuncaoNFe = 0
'    xMovSolicitacaoFuncaoNFe.ChaveAcessoNFe_MovSolicitacaoFuncaoNFe = ""
'    xMovSolicitacaoFuncaoNFe.IPComputadorAC_MovSolicitacaoFuncaoNFe = GetIPAddress()
'    xMovSolicitacaoFuncaoNFe.IPInternetAC_MovSolicitacaoFuncaoNFe = "200??.??.??.??"
'    xMovSolicitacaoFuncaoNFe.SegurancaEstabelecimento_MovSolicitacaoFuncaoNFe = "1234"
'    xMovSolicitacaoFuncaoNFe.CodigoUsuario_MovSolicitacaoFuncaoNFe = g_usuario
'    xMovSolicitacaoFuncaoNFe.VersaoAC_MovSolicitacaoFuncaoNFe = gVersaoSGP
'    xMovSolicitacaoFuncaoNFe.VersaoHost_MovSolicitacaoFuncaoNFe = "??"
'    xMovSolicitacaoFuncaoNFe.Texto_MovSolicitacaoFuncaoNFe = ""
'    xMovSolicitacaoFuncaoNFe.HoraAnalise_MovSolicitacaoFuncaoNFe = CDate("00:00:00")
'    xMovSolicitacaoFuncaoNFe.HoraAprovacao_MovSolicitacaoFuncaoNFe = CDate("00:00:00")
'    xMovSolicitacaoFuncaoNFe.HoraCancelamentoHost_MovSolicitacaoFuncaoNFe = CDate("00:00:00")
'    xMovSolicitacaoFuncaoNFe.HoraConfirmacaoAC_MovSolicitacaoFuncaoNFe = CDate("00:00:00")
'    xMovSolicitacaoFuncaoNFe.HoraCancelamentoAC_MovSolicitacaoFuncaoNFe = CDate("00:00:00")
'    xMovSolicitacaoFuncaoNFe.Mensagem_MovSolicitacaoFuncaoNFe = ""
'    xMovSolicitacaoFuncaoNFe.CodigoRetorno_MovSolicitacaoFuncaoNFe = 0
'    xMovSolicitacaoFuncaoNFe.NumeroLote_MovSolicitacaoFuncaoNFe = 0
'
'    If xMovSolicitacaoFuncaoNFe.Incluir Then
'        GravaSolicitacaoComunicacaoPetromovelAuto = xMovSolicitacaoFuncaoNFe.NSU_MovSolicitacaoFuncaoNFe
'    End If
'
'    Exit Function
'TrataError:
'    Call CriaLogSGP("[GravaSolicitacaoComunicacaoPetromovelAuto]", "Erro ao tentar gravar solicitação de comunicação com PetromovelAuto - Err.Description=" & Err.Description, "")
'    MsgBox "Não foi possível incluir registro de Solicitação de comunicação com PetromovelAuto!", vbCritical, "Erro de Integridade"
'End Function
Private Function LoopIncluiRegistroSolicitacaoNFe() As Boolean
    Dim i As Integer
    LoopIncluiRegistroSolicitacaoNFe = False
        For i = 1 To 30
            If MovSolicitacaoFuncaoNFe.Incluir Then
                Call CriaLogECF(Time & " - [LoopIncluiRegistroSolicitacaoNFe] MovSolicitacaoFuncaoNFe.TipoOperacao_MovSolicitacaoFuncaoNFe=" & MovSolicitacaoFuncaoNFe.TipoOperacao_MovSolicitacaoFuncaoNFe & " - MovSolicitacaoFuncaoNFe.NumeroNFe_MovSolicitacaoFuncaoNFe=" & MovSolicitacaoFuncaoNFe.NumeroNFe_MovSolicitacaoFuncaoNFe)
                LoopIncluiRegistroSolicitacaoNFe = True
                Exit For
            End If
        Next
        If LoopIncluiRegistroSolicitacaoNFe = False Then
            MsgBox "Não foi possível incluir registro de Solicitação de Função NFC-e!", vbCritical, "Erro de Integridade"
        End If
End Function

Private Function RemoveItensNFCeAbandonada(ByVal pCabecalho As CadastroDLL.cMovDocEletronicoCabecalho) As Boolean
    RemoveItensNFCeAbandonada = False

    Dim xMovDocEletronicoItem As New CadastroDLL.cMovDocEletronicoItem
    
    If pCabecalho Is Nothing Then
        MsgBox "Não foi possível excluir itens da NFCe abortada!", vbCritical, "Erro de Integridade"
    ElseIf pCabecalho.CodigoUltimoEvento = Val(EVENTO_NFCE.ABANDONADA) Then
        RemoveItensNFCeAbandonada = xMovDocEletronicoItem.ExcluirItensDocumento(pCabecalho.IdEstabelecimento, pCabecalho.DataEmissao, pCabecalho.Entrada, pCabecalho.Saida, pCabecalho.Modelo, pCabecalho.Serie, lNumeroCupom)
    Else
        MsgBox "Os Itens da NFCe abortada não podem ser exluídos!", vbCritical, "Erro de Integridade"
    End If

End Function
Private Sub btnAbortaNFCe_Click()

    On Error GoTo TrataError
    
        Call AbortarNFCeAtual
    
    Exit Sub
TrataError:
    Call CriaLogSGP("[btnAbortaNFCe_Click]", "Erro ao tentar abortar NFCe: " & Err.Number & " - " & Err.Description, "")
    MsgBox "ERRO: Não foi possível abortar NFCe!" & vbCrLf & Err.Number & " - " & Err.Description, vbCritical, "Erro de Integridade"
End Sub

Private Sub btnDescontoCartaoEspecial_Click()
'Aplica desconto configurado na configuração diversa PERCENTUAL DESCONTO CARTAO ESPECIAL
    Dim xString As String
    If lValorDescontoConcedido = 0 Then
        If fValidaValor2(txt_valor_desconto.Text) = 0 Then
            If ConfiguracaoDiversa.LocalizarCodigo(1, "PERCENTUAL DESCONTO CARTAO ESPECIAL") Then
                If ConfiguracaoDiversa.Valor > 0 Then
                    lValorDescontoConcedido = Round(fValidaValor(txt_valor_recebido.Text) * ConfiguracaoDiversa.Valor / 100, 2)
                End If
            End If
            txt_valor_desconto.Text = Format(lValorDescontoConcedido, "###,##0.00")
            txt_valor_desconto.Enabled = False
            lbl_valor_compra.Caption = Format(lTotalCupom - fValidaValor(txt_valor_desconto.Text), "###,##0.00")
            txt_valor_recebido.Text = Format(lTotalCupom - fValidaValor(txt_valor_desconto.Text), "###,##0.00")
            lDescontoCartaoEspecial = True
            If lImpBematech Then
                Call CriaLogCupom("Desconto Cartão Especial - lValorDescontoConcedido=" & lValorDescontoConcedido)
                'Desconto para o Cupom Fiscal
                xString = Mid(Format(lValorDescontoConcedido, "000000000000.00"), 1, 12) + Mid(Format(lValorDescontoConcedido, "000000000000.00"), 14, 2)
            End If
        Else
            MsgBox "O Valor do desconto deve estar zerado para gerar desconto especial.", vbInformation, "Solicitação Negada!"
        End If
    Else
        MsgBox "Não será permitido gerar desconto novamente na mesma NFCe.", vbInformation, "Solicitação Negada!"
    End If

End Sub

Private Sub cbo_forma_pagamento_GotFocus()
    l_mensagem = Space(165) & "Selecione a forma de pagamento | Tecle F9 para abortar."
    'If l_codigo_cliente = 0 Then
    '    cbo_forma_pagamento.ListIndex = 0
    'Else
    '    cbo_forma_pagamento.ListIndex = 4
    'End If
End Sub

Private Sub cbo_forma_pagamento_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 120 Then
        Call AbortarNFCeAtual
        Exit Sub
    End If

End Sub

Private Sub cbo_forma_pagamento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_valor_recebido.SetFocus
    End If
End Sub
Private Sub cbo_forma_pagamento_LostFocus()
    If cbo_forma_pagamento.ListIndex = -1 Then
        cbo_forma_pagamento.SetFocus
    ElseIf cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 8 Then
        If lOrdem > 2 Then
            MsgBox "Este cupom tem mais de 1 ítem." & Chr(10) & Chr(10) & "Para venda com Ticket Car Smart," & Chr(10) & "será aceito apenas 1 ítem por cupom." & Chr(10) & "Escolha outra forma de pagamento.", vbInformation, "Forma de pagamento não aceita!"
            cbo_forma_pagamento.SetFocus
        End If
    Else
        If cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) < 2 Or cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) > 3 Then
            DespreparaDadosAdicionaisFechamento
        Else
            PreparaDadosAdicionaisFechamento
        End If
    End If
End Sub
Private Sub cboBico_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtQuantidadeDescarregamento.SetFocus
    End If
End Sub
Private Sub cboTipoSubEstoque_GotFocus()
    l_mensagem = Space(165) & "Selecione o tipo do estoque ou Tecle Esc para sair."
    SendMessageLong cboTipoSubEstoque.hWnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cboTipoSubEstoque_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_cliente.SetFocus
    ElseIf KeyAscii = 27 Then
        KeyAscii = 0
        Unload Me
    End If
End Sub
Private Sub chkDesconto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmdOkDescarregar.SetFocus
    End If
End Sub
Private Sub cmd_bico_Click(Index As Integer)
    Dim xLocalErro As Integer
    
    On Error GoTo FileError

    xLocalErro = 1
    
    If lOrdem >= gQUANTIDADE_MAXIMA_ITENS_NFCE Then
       MsgBox "Não é permitido venda com mais de " & gQUANTIDADE_MAXIMA_ITENS_NFCE & " ítens.", vbInformation + vbOKOnly, "Quantidade máxima de itens"
       Exit Sub
    End If
    
    If lAutomacaoStatusBico(Index) = 6 And ValidaCliente And Not ProcessamentoComplementarEmAndamento Then
        If Bomba.LocalizarCodigo(g_empresa, lAutomacaoBico(Index)) Then
            If Bomba.CodigoFisicoBomba = 0 Then
                MsgBox "Não está configurado o código Físico da Bomba." & vbCrLf & "Entre em contato com o suporte Técnico.", vbInformation + vbOKOnly, "Erro de Consistência!"
                Exit Sub
            End If
        End If
        AtivaDesativaBicos (False)
        
        xLocalErro = 100
        Call AutomacaoAlteraArredondamento(lAutomacaoData(Index), lAutomacaoHora(Index), lAutomacaoBico(Index), Index)
        txt_produto.Text = lAutomacaoCodigoProduto(Index)
        txt_quantidade.Text = Format(lAutomacaoLitros(Index), "###,##0.000")
        txt_valor_total.Text = Format(lAutomacaoTotalAPagar(Index), "###,##0.00")
        txt_produto_LostFocus
        txt_quantidade.Text = Format(lAutomacaoLitros(Index), "###,##0.000")
        txt_valor_total.Text = Format(lAutomacaoTotalAPagar(Index), "###,##0.00")
        txt_valor_unitario.Text = Format(lAutomacaoValorLitro(Index), "###,##0.0000")
        xLocalErro = 200
        VerificaDescontoPersonalizado
        xLocalErro = 250
        lAutomacaoBicoEmAcerto = lAutomacaoBico(Index)
        lAutomacaoFlagVendaAutomatica = True
        lAutomacaoBicoEmAcerto = lAutomacaoBico(Index)
        lAutomacaoDataEmAcerto = lAutomacaoData(Index)
        lAutomacaoHoraEmAcerto = lAutomacaoHora(Index)
        lAutomacaoTempoAbastecimentoEmAcerto = "" & lAutomacaoTempoAbastecimento(Index)
        lAutomacaoFormaImpressao = "BOTAO"
        'Manda Imprimir o Cupom Fiscal
        lOrigemAutomacao = True
        xLocalErro = 300
        GravaItem
        xLocalErro = 320
        Call AutomacaoAlteraTabeAbastecimento(Index, lNumeroCupom)
        xLocalErro = 330
        lAutomacaoBicoEmAcerto = 0
        lAutomacaoDataEmAcerto = 0
        lAutomacaoHoraEmAcerto = 0
        lAutomacaoTempoAbastecimentoEmAcerto = ""
        lAutomacaoFormaImpressao = ""
        lAutomacaoBicoEmAcerto = 0
        lAutomacaoFlagVendaAutomatica = False
        xLocalErro = 400
        lAutomacaoStatusBico(Index) = 0
        lbl_automacao_valor(Index).Caption = ""
        lAutomacaoCodigoProduto(Index) = 0
        lAutomacaoBico(Index) = 0
        lAutomacaoData(Index) = 0
        lAutomacaoHora(Index) = 0
        lAutomacaoTempoAbastecimento(Index) = ""
        lAutomacaoValorLitro(Index) = 0
        lAutomacaoLitros(Index) = 0
        lAutomacaoTotalAPagar(Index) = 0
        xLocalErro = 500
        AtivaDesativaBicos (True)
        xLocalErro = 520
    End If
    Exit Sub

FileError:
    Call CriaLogECF(Date & " " & Time & " Index: " & Index & " - xLocalErro: " & xLocalErro & " cmd_bico_Click - Erro=" & Err.Number & " - " & Err.Description)
    MsgBox "Erro não identificado. cmd_bico_Click. Bico: " & lAutomacaoBico(Index), vbCritical, "Erro: cmd_bico_Click"
    
End Sub
Private Sub cmd_bico_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Shift = 2 And ValidaCliente Then
        If UCase(g_cidade_empresa) Like "*REDEN*" Or UCase(g_cidade_empresa) Like "*CUMAR*" Or UCase(g_cidade_empresa) Like "*CONCEI*" Then
            If fValidaValor(lbl_automacao_valor(Index).Caption) > 0 Then
                Call GeraDescontoAbastecimento(lAutomacaoBico(Index), lAutomacaoData(Index), lAutomacaoHora(Index), lAutomacaoValorLitro(Index), lAutomacaoTotalAPagar(Index))
            End If
        End If
    End If
End Sub
Private Sub cmd_cancelar_ponto_Click()
    Unload Me
End Sub
Private Sub cmd_cancelar2_Click()
'    'frm_fechamento_cupom.Width = 5700
'    'frm_fechamento_cupom.Top = 500
'    'frm_fechamento_cupom.Left = 120
'    frm_fechamento_cupom.Top = 200
'    frm_fechamento_cupom.Left = 200
'    frm_fechamento_cupom.Height = 1
'    frm_fechamento_cupom.Width = 1
'    frm_fechamento_cupom.ZOrder 1
'    frm_fechamento_cupom.Enabled = False
'    cmdIniciaProcessoFinalizacaoNFCe.Visible = True
'    frmDados.Enabled = True
'    NovoCupom
    
    Call CancelarFinalizacaoNFCe
    
End Sub
Private Sub cmd_cancelar2_GotFocus()
    l_mensagem = Space(165) & "Tecle enter para informar mais produto."
End Sub
Private Sub cmd_DescontoCorreio_Click()
    Dim xString As String
    If lValorDescontoConcedido = 0 Then
        If fValidaValor2(txt_valor_desconto.Text) = 0 Then
        ''    ChamaCartaoDesconto
            'txt_valor_desconto.Text = Format(fValidaValor(txt_valor_recebido.Text) * 3 / 100, "###,##0.00")
            'lValorDescontoConcedido = Round(fValidaValor(txt_valor_recebido.Text) * 3.5 / 100, 2)
            If ConfiguracaoDiversa.LocalizarCodigo(1, "PERCENTUAL CARTAO DESCONTO CORREIOS") Then
                If ConfiguracaoDiversa.Valor > 0 Then
                    lValorDescontoConcedido = Round(fValidaValor(txt_valor_recebido.Text) * ConfiguracaoDiversa.Valor / 100, 2)
                End If
            End If
            txt_valor_desconto.Text = Format(lValorDescontoConcedido, "###,##0.00")
            txt_valor_desconto.Enabled = False
            lbl_valor_compra.Caption = Format(lTotalCupom - fValidaValor(txt_valor_desconto.Text), "###,##0.00")
            txt_valor_recebido.Text = Format(lTotalCupom - fValidaValor(txt_valor_desconto.Text), "###,##0.00")
            If lImpBematech Then
                Call CriaLogCupom("Cartão Desconto Correios - lValorDescontoConcedido=" & lValorDescontoConcedido)
                'Desconto para o Cupom Fiscal
                xString = Mid(Format(lValorDescontoConcedido, "000000000000.00"), 1, 12) + Mid(Format(lValorDescontoConcedido, "000000000000.00"), 14, 2)
                'Call CriaLogCupom("Bematech_FI_IniciaFechamentoCupom('D', '$', xString) xString=" & xString)
                'BemaRetorno = Bematech_FI_IniciaFechamentoCupom("D", "$", xString)
                'Call CriaLogCupom("Bematech_FI_IniciaFechamentoCupom - BemaRetorno=" & BemaRetorno)
            End If
        Else
            MsgBox "O Valor do desconto deve estar zerado para gerar desconto Correios.", vbInformation, "Solicitação Negada!"
        End If
    Else
        MsgBox "Não será permitido gerar desconto Correios novamente no mesmo Cupom Fiscal.", vbInformation, "Solicitação Negada!"
    End If
End Sub
Private Sub cmd_fila_ok_Click()
    Dim i As Integer
    
    If lOrdem >= gQUANTIDADE_MAXIMA_ITENS_NFCE Then
       MsgBox "Não é permitido venda com mais de " & gQUANTIDADE_MAXIMA_ITENS_NFCE & " ítens.", vbInformation + vbOKOnly, "Quantidade máxima de itens"
       Exit Sub
    End If
    
    If MSFlexGrid.TextMatrix(MSFlexGrid.Row, 0) <> "" Then
        Call AutomacaoAlteraArredondamento(CDate(MSFlexGrid.TextMatrix(MSFlexGrid.Row, 3)), CDate(MSFlexGrid.TextMatrix(MSFlexGrid.Row, 4)), Val(MSFlexGrid.TextMatrix(MSFlexGrid.Row, 0)), -1)
        If MovimentoAbastecimento.LocalizarCodigo(g_empresa, CDate(MSFlexGrid.TextMatrix(MSFlexGrid.Row, 3)), CDate(MSFlexGrid.TextMatrix(MSFlexGrid.Row, 4)), Val(MSFlexGrid.TextMatrix(MSFlexGrid.Row, 0))) Then
            i = MovimentoAbastecimento.Bico - 1
            txt_produto.Text = MovimentoAbastecimento.CodigoProduto
            txt_quantidade.Text = MovimentoAbastecimento.Quantidade
            txt_valor_total.Text = MovimentoAbastecimento.ValorTotal - MovimentoAbastecimento.ValorDesconto
            txt_produto_LostFocus
            txt_quantidade.Text = Format(MovimentoAbastecimento.Quantidade, "###,##0.000")
            txt_valor_total.Text = Format(MovimentoAbastecimento.ValorTotal - MovimentoAbastecimento.ValorDesconto, "###,##0.00")
            txt_valor_unitario.Text = Format(MovimentoAbastecimento.ValorUnitario, "###,##0.0000")
            VerificaDescontoPersonalizado
            lAutomacaoBicoEmAcerto = MovimentoAbastecimento.Bico
            lAutomacaoDataEmAcerto = MovimentoAbastecimento.Data
            lAutomacaoHoraEmAcerto = MovimentoAbastecimento.Hora
            lAutomacaoTempoAbastecimentoEmAcerto = "" & MovimentoAbastecimento.TempoAbastecimento
            lAutomacaoFormaImpressao = "FILA"
            lAutomacaoBicoEmAcerto = MovimentoAbastecimento.Bico
            lAutomacaoData(i) = MovimentoAbastecimento.Data
            lAutomacaoHora(i) = MovimentoAbastecimento.Hora
            lAutomacaoBico(i) = MovimentoAbastecimento.Bico
            lAutomacaoFlagVendaAutomatica = True
            'Manda Imprimir o Cupom Fiscal
            lOrigemAutomacao = True
            GravaItem
            Call AutomacaoAlteraTabeAbastecimento(i, lNumeroCupom)
            lAutomacaoBicoEmAcerto = 0
            lAutomacaoDataEmAcerto = 0
            lAutomacaoHoraEmAcerto = 0
            lAutomacaoTempoAbastecimentoEmAcerto = ""
            lAutomacaoFormaImpressao = ""
            lAutomacaoBicoEmAcerto = 0
            lAutomacaoFlagVendaAutomatica = False
            lAutomacaoStatusBico(i) = 0
            lbl_automacao_valor(i).Caption = ""
            lAutomacaoCodigoProduto(i) = 0
            lAutomacaoBico(i) = 0
            lAutomacaoData(i) = 0
            lAutomacaoHora(i) = 0
            lAutomacaoTempoAbastecimento(i) = ""
            lAutomacaoValorLitro(i) = 0
            lAutomacaoLitros(i) = 0
            lAutomacaoTotalAPagar(i) = 0
        Else
            MsgBox "Não foi possível localizar o abastecimento!", vbInformation, "Erro de Integridade!"
        End If
    End If
    frm_fila_bico.Enabled = False
    frm_fila_bico.Visible = False
    txt_produto.SetFocus
End Sub
Private Sub cmd_fila_sair_Click()
    frm_fila_bico.Enabled = False
    frm_fila_bico.Visible = False
    txt_produto.SetFocus
End Sub
Private Sub cmd_ok_ponto_Click()
    Dim xChamaCaixa As Boolean 'new 17/08/2015
    
    xChamaCaixa = False 'new 17/08/2015
    
    BuscaPeriodo
    
    If ValidaCamposPonto Then
        If lImpQuick Then
            If EcfQuickSemPapel Then
                MsgBox "ECF não está em linha ou sem papel.", vbInformation, "Erro na ECF!"
                txt_funcionario_ponto.SetFocus
                Exit Sub
            End If
        
            'teste para fechar gerencial caso esteja aberto
            If EcfQuickLeRegistrador("EstadoFiscal", "Long", 5) = "32" Then
                Call EcfQuickEncerraDocumento(0, "Gerencial")
            End If
            If EcfQuickLeRegistrador("EstadoFiscal", "Long", 5) = "64" Then
                Call EcfQuickEncerraDocumento(0, "Gerencial")
            End If
            
            If EcfQuickLeRegistrador("EstadoFiscal", "Long", 5) = "2" Then
                Dim xValor As Currency
                'Cancela o cupom aberto
                Call EcfQuickCancelaCupom
                mnuSenha_Click
            End If
        End If
        If txt_senha_ponto.Text <> "" Then
            txt_senha_ponto.Text = Kriptografa(txt_senha_ponto.Text)
            If txt_senha_ponto.Text = l_senha_funcionario Then
                If lCaixaIndividual Then
                    l_codigo_funcionario = Val(dtcboFuncionario.BoundText)
                    'verifica se o tipo de movimento do funcionario é troca de oleo = 3
                    'ou se é pista(frentista) = 2
                    If PeriodoTrocaOleo.LocalizarCodigo(g_empresa, Val(txt_funcionario_ponto.Text)) Then
                        lTipoMovimento = 3
                    Else
                        lTipoMovimento = 2
                    End If
                    
                    If ExisteCaixaIndividualAberto(Date) = False Then
                        Exit Sub
                    End If
                End If
                Call AbilitaMenu(True)
                g_usuario = Usuario.Codigo
                g_nome_usuario = Usuario.Nome
                g_nivel_acesso = Usuario.TipoAcesso
                lValorDescontoConcedido = 0
                
                'Zerar cliente ao abrir uma nova nfce
                l_codigo_cliente = 0
                txt_cliente.Text = "0"
                dtcboCliente.BoundText = 0
                
                lDescontoPostoAki = False
                lIntegraDescontoCartaoCorreios = False
                lIntegraDescontoCartaoEspecial = False
                
                lRestringeDescontoNivelUsuario = False
                If ConfiguracaoDiversa.LocalizarCodigo(1, "NFCEAUTO: RESTRINGE CAMPO DESCONTO") Then
                    
                    If g_nivel_acesso >= ConfiguracaoDiversa.Codigo Then
                        lRestringeDescontoNivelUsuario = True
                    End If
                    
                End If

                
                If lRestringeDescontoNivelUsuario Then
                    txt_valor_desconto.Enabled = False
                Else
                    txt_valor_desconto.Enabled = True
                End If
                
                menu_personalizado.StatusBar1.Panels(2).Text = g_nome_usuario
                menu_personalizado.StatusBar1.Panels(2).AutoSize = sbrContents
                l_codigo_funcionario = Val(dtcboFuncionario.BoundText)
                l_nome_funcionario = dtcboFuncionario
                Me.Caption = "Nota Fiscal do Consumidor Eletrônica - " & l_nome_funcionario
                frm_ponto.Enabled = False
                frm_ponto.ZOrder 1
'                mnuReducaoZ.Visible = True
'                If g_nivel_acesso > 3 Then
'                    mnuReducaoZ.Enabled = False
'                Else
'                    mnuReducaoZ.Enabled = True
'                End If
               ' AtivaReducaoZ
                frmDados.Enabled = True
                txt_cupom_fiscal.Enabled = True
                NovoCupom
                AtualizaBombasAbastecimento
                AutomacaoMostraBicos
               

                
                If lData > g_cfg_data_i Then
                    If (MsgBox("A data do computador está maior que a data de operação do caixa." & vbLf & "Data do computador: " & Format(lData, "dd/mm/yyyy") & vbLf & "Data do Caixa: " & Format(g_cfg_data_i, "dd/mm/yyyy") & vbLf & "Deseja mudar a data de liberação do caixa?", vbQuestion + vbYesNo + vbDefaultButton1, "Muda data de Abertura do Caixa!")) = vbYes Then
                        If LiberacaoDigitacao.LocalizarCodigo(g_empresa, lTipoMovimento) Then
                            Call GravaAuditoria(1, Me.name, 10, "De: Data:" & fMascaraData(LiberacaoDigitacao.DataInicial) & " - " & " Per:" & LiberacaoDigitacao.PeriodoInicial & " Tp.Mov:" & LiberacaoDigitacao.TipoMovimento)
                            LiberacaoDigitacao.DataInicial = lData
                            LiberacaoDigitacao.DataFinal = lData
                            LiberacaoDigitacao.PeriodoInicial = 1
                            LiberacaoDigitacao.PeriodoFinal = 1
                            Call GravaAuditoria(1, Me.name, 10, "Para: Data:" & fMascaraData(LiberacaoDigitacao.DataInicial) & " - " & " Per:" & LiberacaoDigitacao.PeriodoInicial & " Tp.Mov:" & LiberacaoDigitacao.TipoMovimento)
                            If LiberacaoDigitacao.Alterar(g_empresa, lTipoMovimento) Then
                                g_cfg_data_i = LiberacaoDigitacao.DataInicial
                                g_cfg_data_f = LiberacaoDigitacao.DataInicial
                                g_cfg_periodo_i = LiberacaoDigitacao.PeriodoInicial
                                g_cfg_periodo_f = LiberacaoDigitacao.PeriodoInicial
                                lPeriodo = g_cfg_periodo_i
                            Else
                                MsgBox "Não foi possível alterar liberação de digitação!", vbInformation, "Erro Interno!"
                            End If
                        Else
                            MsgBox "Não foi possível localizar registro de liberação de digitação!", vbInformation, "Erro de Verificação!"
                        End If
                    End If
                End If
            Else
                MsgBox "Senha informada não confere." & Chr(10) & "Informe pela " & 2 & "a vez.", vbInformation, "Senha Inválida!"
                txt_senha_ponto.Text = ""
                txt_senha_ponto.SetFocus
            End If
        End If
    End If
End Sub

Private Sub cmdAtualizaAbastecimento_Click()
    cmdAtualizaAbastecimento.Enabled = False
    Call ConsisteNFCePendentes(2)
    Call AguardaMS(3000)
    AtualizaBombasAbastecimento
    AutomacaoMostraBicos
    cmdAtualizaAbastecimento.Enabled = True
End Sub

Private Sub cmdDescontoAppPostoAki_Click()
    If lValorDescontoConcedido = 0 Then
        If gUfEmpresa = "PA" Then
            If fValidaValor2(txt_valor_desconto.Text) = 0 Then
                If (MsgBox("Deseja aplicar desconto?" & vbCrLf & "Escolha SIM - Para Desconto." & vbCrLf & "Escolha NÃO - Para Pontuação.", vbQuestion + vbYesNo + vbDefaultButton1, "Desconto ou Pontuação")) = vbYes Then
                    If (MsgBox("Esta operação é uma troca de Óleo?" & vbCrLf & "Escolha NÃO - Para Desconto em Combustível." & vbCrLf & "Escolha SIM - Para Desconto em Troca de Óleo.", vbQuestion + vbYesNo + vbDefaultButton2, "Tipo de Desconto")) = vbYes Then
                        Call ChamaCartaoDesconto("POSTOAKI", "", True, False)
                    Else
                        Call ChamaCartaoDesconto("POSTOAKI", "", False, False)
                    End If
                Else
                    Call ChamaCartaoDesconto("POSTOAKI", "", False, True)
                End If
            Else
                MsgBox "O Valor do desconto deve estar zerado para solicitar Desconto pelo App PostoAki.", vbInformation, "Solicitação Negada!"
            End If
        Else
            If fValidaValor2(txt_valor_desconto.Text) = 0 Then
                If (MsgBox("Esta operação é uma troca de Óleo?" & vbCrLf & "Escolha NÃO - Para Desconto em Combustível." & vbCrLf & "Escolha SIM - Para Desconto em Troca de Óleo.", vbQuestion + vbYesNo + vbDefaultButton2, "Tipo de Desconto")) = vbYes Then
                    Call ChamaCartaoDesconto("POSTOAKI", "", True, False)
                Else
                    Call ChamaCartaoDesconto("POSTOAKI", "", False, False)
                End If
            Else
                MsgBox "O Valor do desconto deve estar zerado para solicitar Desconto pelo App PostoAki.", vbInformation, "Solicitação Negada!"
            End If
        End If
    Else
        MsgBox "Não será permitido solicitar desconto pelo App PostoAki novamente no mesmo Cupom Fiscal.", vbInformation, "Solicitação Negada!"
    End If
End Sub
Private Sub cmdIniciaProcessoFinalizacaoNFCe_Click()
    If l_flag_cupom_fiscal = "A" Then
        Call GravaAuditoria(1, Me.name, 23, cmdIniciaProcessoFinalizacaoNFCe.ToolTipText & " Func.:" & l_nome_funcionario)
        lInformaFormaPagamento = True
        Call CriaLogMonitoraBug("cmdIniciaProcessoFinalizacaoNFCe_Click", "Movimento_NfceAuto", "Chamou CancelaCupom")
        CancelaCupom
    End If
End Sub
Private Sub cmdIniciaProcessoNFCe_Click()
    Call CriaLogMonitoraBug("cmdIniciaProcessoNFCe_Click", "Movimento_NfceAuto", "Chamou IniciaProcessamentoNFCe")
    cmdIniciaProcessoNFCe.Enabled = False
    
    IniciaProcessamentoNFCe
End Sub

Private Sub AtualizaBaseCalculoICMSCabecalho()

Dim ResultadoCalculo As New ADODB.Recordset


    lSQL = "SELECT DataEmissao_MovDEItem, Modelo_MovDEItem, Serie_MovDEItem, Numero_MovDEItem,"
    lSQL = lSQL & " SUM(ValorBcIcms_MovDEItem) As BcIcms, SUM(ValorBcIcmsSt_MovDEItem) As BcIcmsSt"
    lSQL = lSQL & " FROM MovimentoDocumentoEletronicoItem "
    lSQL = lSQL & " WHERE IdEstabelecimento_MovDEItem = " & g_empresa
    lSQL = lSQL & " AND DataEmissao_MovDEItem = " & preparaData(lDataCupom)
    lSQL = lSQL & " AND Numero_MovDEItem = " & preparaTexto(lNumeroCupom)
    lSQL = lSQL & " AND Saida_MovDEItem = " & preparaBooleano(True)
    lSQL = lSQL & " GROUP BY DataEmissao_MovDEItem, Modelo_MovDEItem, Serie_MovDEItem, Numero_MovDEItem"
    lSQL = lSQL & " ORDER BY DataEmissao_MovDEItem, Modelo_MovDEItem, Serie_MovDEItem, Numero_MovDEItem"


Set ResultadoCalculo = Conectar.RsConexao(lSQL)

If ResultadoCalculo.RecordCount > 0 Then

    
    MovDocEletronicoCabecalho.ValorBCICMS = ResultadoCalculo!BcIcms - MovDocEletronicoCabecalho.ValorDesconto
    MovDocEletronicoCabecalho.ValorBCICMSST = ResultadoCalculo!BcIcmsSt - MovDocEletronicoCabecalho.ValorDesconto
End If


End Sub

Private Sub cmdIniciaProcessoNFCe_GotFocus()
    l_mensagem = Space(165) & "Tecle enter para finalizar o cumpo fiscal."
End Sub
Private Sub cmd_abastecimentos_nao_recebidos_Click()
    Dim i As Integer
    Dim xFaseErro As Integer
    Dim xSQL As String
    
    On Error GoTo FileError
    
    xFaseErro = 1
    frm_fila_bico.Enabled = True
    xFaseErro = 2
    frm_fila_bico.Left = 80
    xFaseErro = 3
    frm_fila_bico.Top = 780
    xFaseErro = 4
    frm_fila_bico.Visible = True
    xFaseErro = 5
    frm_fila_bico.ZOrder 0
    xFaseErro = 6
    lbl_fila_valor.Caption = ""
    xFaseErro = 7
    lbl_fila_litros.Caption = ""
    xFaseErro = 8
    lbl_fila_total.Caption = ""
    
    xFaseErro = 9
    xSQL = "SELECT Movimento_Abastecimento.Bico, Produto.Nome, (Movimento_Abastecimento.[Valor Total] - Movimento_Abastecimento.[Valor do Desconto]) AS [Valor Total], Movimento_Abastecimento.Data, Movimento_Abastecimento.Hora"
    xSQL = xSQL & " FROM  Movimento_Abastecimento, Produto"
    xSQL = xSQL & " WHERE Movimento_Abastecimento.Acerto = " & preparaBooleano(False)
    xSQL = xSQL & " AND   Movimento_Abastecimento.[Codigo do Produto] = Produto.Codigo"
    If lCaixaIndividual Then
        xSQL = xSQL & " AND ( Movimento_Abastecimento.[Codigo do Funcionario] = " & l_codigo_funcionario
        xSQL = xSQL & " OR Movimento_Abastecimento.[Tempo de Abastecimento] = " & preparaTexto("11111") & " )"
    End If
    If (MsgBox("Escolha SIM - Para ordernar abastecimentos pelo Valor." & vbCrLf & "Escolha NÃO - Para ordernar abastecimentos pelo Código da Bomba.", vbQuestion + vbYesNo + vbDefaultButton2, "Abastecimentos NÃO Recebidos")) = vbYes Then
        xFaseErro = 10
        xSQL = xSQL & " ORDER BY Movimento_Abastecimento.[Valor Total], Movimento_Abastecimento.Data, Movimento_Abastecimento.Hora"
    Else
        xFaseErro = 11
        xSQL = xSQL & " ORDER BY Movimento_Abastecimento.Bico, Movimento_Abastecimento.Data, Movimento_Abastecimento.Hora"
    '    xSQL = xSQL & " ORDER BY Data DESC, Hora DESC, [Valor Total] DESC"
    End If
    xFaseErro = 12
    LimpaMSFlexGrid
    'Abre RecordSet
    xFaseErro = 3
    Set rsTabela = New ADODB.Recordset
    Set rsTabela = Conectar.RsConexao(xSQL)
    'Verifica movimento
    i = 0
    xFaseErro = 14
    MSFlexGrid.Visible = False
    xFaseErro = 15
    If rsTabela.RecordCount > 0 Then
        xFaseErro = 16
        Do Until rsTabela.EOF
            xFaseErro = 17
            MSFlexGrid.Rows = MSFlexGrid.Rows + 1
            i = i + 1
            MSFlexGrid.Row = i
            MSFlexGrid.Col = 0
            MSFlexGrid.Text = rsTabela!Bico
            MSFlexGrid.Col = 1
            MSFlexGrid.Text = rsTabela!Nome
            MSFlexGrid.Col = 2
            MSFlexGrid.Text = Format(rsTabela![Valor Total], "##,###,##0.00")
            MSFlexGrid.Col = 3
            MSFlexGrid.Text = Format(rsTabela!Data, "dd/MM/yyyy")
            MSFlexGrid.Col = 4
            MSFlexGrid.Text = Format(rsTabela!Hora, "HH:mm:ss")
            rsTabela.MoveNext
        Loop
    End If
    xFaseErro = 18
    If rsTabela.State = 1 Then
        rsTabela.Close
    End If
    xFaseErro = 19
    MSFlexGrid.Visible = True
    xFaseErro = 20
    MSFlexGrid.SetFocus
    xFaseErro = 21
    Exit Sub

FileError:
    Call CriaLogECF(Date & " " & Time & " cmd_abastecimentos_nao_recebidos_Click: xFaseErro=" & xFaseErro & " - Erro=" & Err.Number & " - " & Err.Description)
    Call CriaLogECF(Date & " " & Time & " IntegraCartaoCreditoNoCaixa: xSQL=" & xSQL)
    MsgBox "Erro não identificado. cmd_abastecimentos_nao_recebidos_Click.", vbCritical, "Erro: cmd_abastecimentos_nao_recebidos_Click"
End Sub
Private Sub cmdCancelarDescarregar_Click()
    frmDescarregar.Enabled = False
    frmDescarregar.Visible = False
    cmd_bico(0).SetFocus
End Sub
Private Sub cmdCartaoFidelidadeDesconto_Click()
    If lValorDescontoConcedido = 0 Then
        If fValidaValor2(txt_valor_desconto.Text) = 0 Then
            Call ChamaCartaoDesconto("", "", False, False)
        Else
            MsgBox "O Valor do desconto deve estar zerado para solicitar Cartão Fidelidade/Desconto.", vbInformation, "Solicitação Negada!"
        End If
    Else
        MsgBox "Não será permitido solicitar Cartão Fidelidade/Desconto novamente no mesmo Cupom Fiscal.", vbInformation, "Solicitação Negada!"
    End If
End Sub
Private Sub cmdDescarregar_Click()
    Call GravaAuditoria(1, Me.name, 23, cmdDescarregar.ToolTipText & " Func.:" & l_nome_funcionario)
    frmDescarregar.Enabled = True
    frmDescarregar.Visible = True
    cboBico.ListIndex = -1
    txtQuantidadeDescarregamento.Text = ""
    cboBico.SetFocus
End Sub
Private Sub cmdInformaPlacaVeiculo_Click()
    
    On Error GoTo trata_erro
    
    g_string = ""
    InformaPlacaKM.Show 1
    If Len(g_string) > 0 Then
        If RetiraGString(1) <> "" Then
            lPlacaLetra = RetiraGString(1)
        End If
        If RetiraGString(2) <> "" Then
            lPlacaNumero = Val(RetiraGString(2))
        End If
        If RetiraGString(3) <> "" Then
            lKMVeiculo = CLng(RetiraGString(3))
        End If
'        If Trim(txt_observacao_2.Text) = "" Then
'            txt_observacao_2.Text = "Placa: " & lPlacaLetra & "-" & lPlacaNumero & " KM: " & lKMVeiculo
'        Else
            txt_placa.Text = lPlacaLetra & "-"
            If lPlacaNumero > 0 Then
                txt_placa.Text = txt_placa.Text & Format(lPlacaNumero, "0000")
            End If
            txt_kilometragem.Text = lKMVeiculo
'        End If
    End If
    g_string = ""
    If cmdIniciaProcessoNFCe.Enabled Then
        cmdIniciaProcessoNFCe.SetFocus
    End If
    
    Exit Sub

trata_erro:
    MsgBox "Erro na informação da placa do veículo.", vbCritical, "Erro na Placa!"
    g_string = ""
    lPlacaLetra = ""
    lPlacaNumero = 0
    lKMVeiculo = 0
    txt_observacao_2.Text = ""
End Sub
Private Sub cmdOkDescarregar_Click()
    If cboBico.ListIndex = -1 Then
        MsgBox "Selecione um bico a ser descarregado.", vbInformation, "Atenção!"
        cboBico.SetFocus
        Exit Sub
    End If
    If Val(txtQuantidadeDescarregamento.Text) = 0 Then
        MsgBox "Informe a quantidade de abastecimento a ser descarregado.", vbInformation, "Atenção!"
        txtQuantidadeDescarregamento.SetFocus
        Exit Sub
    End If
    DescarregaAbastecimento
    frmDescarregar.Enabled = False
    frmDescarregar.Visible = False
    cmd_bico(0).SetFocus
End Sub
Private Sub ImpValeTroco()
    Dim xString As String
    Dim xNumeroCupom As String
    Dim xValor As String
    Dim i As Integer
    
    'Busca Número do ECF
    xNumeroCupom = Space(6)
    'Call CriaLogCupom("Bematech_FI_NumeroCupom(xNumeroCupom)")
    'BemaRetorno = Bematech_FI_NumeroCupom(xNumeroCupom)
    'Call CriaLogCupom("Bematech_FI_NumeroCupom - xNumeroCupom=" & xNumeroCupom & " - BemaRetorno=" & BemaRetorno)
    xNumeroCupom = "000000"
    xNumeroCupom = Format(Val(xNumeroCupom) + 1, "000000")
    
    i = Len(txt_valor_total.Text)
    
    'Abre o cupom fiscal
    'Call CriaLogCupom("Bematech_FI_AbreCupom('')")
    'BemaRetorno = Bematech_FI_AbreCupom("")
    'Call CriaLogCupom("Bematech_FI_AbreCupom - BemaRetorno=" & BemaRetorno)
    
    
    
    
    'Imprime Produto
    '          123456789012345678              90123456789012345678
    xString = "VALE ABAST. N. " & xNumeroCupom & "   R$            "
    i = Len(txt_valor_total.Text)
    Mid(xString, 27 + i, i) = txt_valor_total.Text
    
    If Val(txt_cliente.Text) > 0 Then
        xString = xString & Format(Me.txt_cliente.Text, "000") & "                                             "
        Mid(xString, 43, 40) = dtcboCliente.Text
    Else
        xString = xString & "          ____________________________          "
    End If
    xString = xString & "Responsável: " & Format(txt_funcionario_ponto.Text, "000") & "                                "
    Mid(xString, 104, 31) = dtcboFuncionario.Text
    xString = xString & "                                                "
    xString = xString & "CLIENTE (RECEBIDO)"
    'Call CriaLogCupom("Bematech_FI_VendeItemDepartamento(... xString ...) xString=" & xString)
    'BemaRetorno = Bematech_FI_VendeItemDepartamento(Format(8888, "#,##0"), xString, "II", "000000010", "0001000", "0000000000", "0000000000", "05", "PO")
    'Call CriaLogCupom("Bematech_FI_VendeItemDepartamento(... xString ...) - xString=" & xString & " - BemaRetorno=" & BemaRetorno)
    
    'Cancela o cupom fiscal
    'Call CriaLogCupom("Bematech_FI_CancelaCupom")
    'BemaRetorno = Bematech_FI_CancelaCupom
    'Call CriaLogCupom("Bematech_FI_CancelaCupom - BemaRetorno=" & BemaRetorno)
    
    
    'Abre o cupom fiscal
    'Call CriaLogCupom("Bematech_FI_AbreCupom")
    'BemaRetorno = Bematech_FI_AbreCupom("")
    'Call CriaLogCupom("Bematech_FI_AbreCupom - BemaRetorno=" & BemaRetorno)
    
    'Imprime Produto
    Mid(xString, 183, 18) = "CAIXA (EMITIDO)   "
    'Call CriaLogCupom("Bematech_FI_VendeItemDepartamento(... xString ...) xString=" & xString)
    'BemaRetorno = Bematech_FI_VendeItemDepartamento(Format(8888, "#,##0"), xString, "II", "000000010", "0001000", "0000000000", "0000000000", "05", "PO")
    'Call CriaLogCupom("Bematech_FI_VendeItemDepartamento(... xString ...) - xString=" & xString & " - BemaRetorno=" & BemaRetorno)
    
    'Cancela o cupom fiscal
    'Call CriaLogCupom("Bematech_FI_CancelaCupom")
    'BemaRetorno = Bematech_FI_CancelaCupom
    'Call CriaLogCupom("Bematech_FI_CancelaCupom - BemaRetorno=" & BemaRetorno)
End Sub
Function IncluiMovimentoCaixa(ByVal pData As Date, ByVal pPeriodo As Integer, ByVal pDesconto As Boolean, ByVal pTipoLancamentoPadrao As String, ByVal pValor As Currency, ByVal pComplementoDadosInterno As String, ByVal pComplementoCaixa As String) As Boolean
    Dim xComplemento As String
    Dim xValorDesconto As Currency
    Dim xContaDebito As String
    Dim xContaCredito As String
    Dim xValor As Currency
    
    IncluiMovimentoCaixa = False
    xValorDesconto = 0
    xValor = 0
    If pTipoLancamentoPadrao = "NotaAbastecimento" Then
        If pDesconto Then
            xValorDesconto = Format(MovNotaAbastecimento.ValorDescontoUnitario * MovNotaAbastecimento.Quantidade, "00000000.00")
            If xValorDesconto > 0 Then
                xComplemento = "NOTA ABASTECIMENTO DESCONTO"
            Else
                xComplemento = "NOTA ABASTECIMENTO ACRESCIMO"
            End If
        Else
            xComplemento = "NOTA ABASTECIMENTO"
            xValorDesconto = DescontoPersonalizado(MovCupomFiscal.CodigoCliente, Cliente.CodigoGrupoCliente, MovCupomFiscal.CodigoProduto, MovCupomFiscal.ValorUnitario)
            If xValorDesconto > 0 Then
                xValorDesconto = Format(xValorDesconto * fValidaValor(MovCupomFiscal.Quantidade), "00000000.00")
            End If
        End If
    ElseIf pTipoLancamentoPadrao = "VENDA DE LUBRIFICANTES" Then
        If IntegracaoCaixa.LocalizarNome(g_empresa, pTipoLancamentoPadrao) Then
            If Val(cboTipoSubEstoque.Text) = 3 Then
                lTipoMovimento = 3
            End If
            xComplemento = "LUBRIFICANTES Per:" & MovCupomFiscal.Periodo & " Ilha:" & lIlha & " S.Est:" & Val(cboTipoSubEstoque.Text) & " T.Mov:" & lTipoMovimento
            'Caso Exista Deleta e Guarda o Valor
            If lCaixaIndividual Then
                If MovCaixaPista.LocalizarRegistroEspecialUsu(g_empresa, MovCupomFiscal.Data, Val(MovCupomFiscal.Periodo), 1, xComplemento, IntegracaoCaixa.ContaCredito, "C", g_usuario) Then
                    xValor = MovCaixaPista.Valor
                    If Not MovCaixaPista.Excluir(g_empresa, MovCupomFiscal.Data, MovCaixaPista.NumeroMovimento) Then
                        MsgBox "Não foi possível excluir o movimento caixa!", vbOKOnly + vbInformation, "Erro de Integridade"
                    End If
                End If
            Else
                If MovCaixaPista.LocalizarRegistroEspecial(g_empresa, MovCupomFiscal.Data, Val(MovCupomFiscal.Periodo), 1, xComplemento, IntegracaoCaixa.ContaCredito, "C") Then
                    xValor = MovCaixaPista.Valor
                    If Not MovCaixaPista.Excluir(g_empresa, MovCupomFiscal.Data, MovCaixaPista.NumeroMovimento) Then
                        MsgBox "Não foi possível excluir o movimento caixa!", vbOKOnly + vbInformation, "Erro de Integridade"
                    End If
                End If
            End If
            xValor = xValor + MovCupomFiscal.ValorTotal
        Else
            MsgBox "Não existe a integração=" & "VENDA DE LUBRIFICANTES" & ".", vbInformation, "Registro Inexistente"
            Exit Function
        End If
        xComplemento = pTipoLancamentoPadrao
    ElseIf pTipoLancamentoPadrao = "CartaoCredito" Then
        xComplemento = "CARTAO " & CartaoCredito.Nome
    ElseIf pTipoLancamentoPadrao = "DESCONTO AUTORIZADO" Then
        If IntegracaoCaixa.LocalizarNome(g_empresa, pTipoLancamentoPadrao) Then
            xComplemento = pTipoLancamentoPadrao
        Else
            MsgBox "Não existe a integração=" & pTipoLancamentoPadrao & ".", vbInformation, "Registro Inexistente"
            Exit Function
        End If
    ElseIf pTipoLancamentoPadrao = "DescontoCartaoFidelidade" Then
        xValorDesconto = pValor
        xComplemento = "DESCONTO CARTÃO FIDELIDADE"
    ElseIf pTipoLancamentoPadrao = "DescontoCartaoCorreios" Then
        xValorDesconto = pValor
        xComplemento = "DESCONTO CARTÃO CORREIOS"
    ElseIf pTipoLancamentoPadrao = "DescontoCartaoEspecial" Then
        xValorDesconto = pValor
        xComplemento = "DESCONTO CARTÃO ESPECIAL"
    Else
        xComplemento = pTipoLancamentoPadrao
    End If
    
    If IntegracaoCaixa.LocalizarNome(g_empresa, xComplemento) Then
        xContaDebito = IntegracaoCaixa.ContaDebito
        xContaCredito = IntegracaoCaixa.ContaCredito
        If pTipoLancamentoPadrao = "NotaAbastecimento" Then
            If pDesconto = False Then
                MovCaixaPista.Valor = MovCupomFiscal.ValorTotal
            Else
                If xValorDesconto > 0 Then
                    MovCaixaPista.Valor = xValorDesconto
                Else
                    MovCaixaPista.Valor = xValorDesconto * -1
                    xContaDebito = IntegracaoCaixa.ContaCredito
                    xContaCredito = IntegracaoCaixa.ContaDebito
                End If
            End If
            xComplemento = Cliente.RazaoSocial
            MovCaixaPista.CodigoHistorico = IntegracaoCaixa.HistoricoPadrao
            MovCaixaPista.DadosInterno = "NOTAA|@|" & MovCupomFiscal.CodigoCliente & "|@|" & MovCupomFiscal.CodigoProduto & "|@|" & MovCupomFiscal.Ordem & "|@|"
            MovCaixaPista.CodigoLancamentoPadrao = 3
            'aqui numero da nota
            MovCaixaPista.NumeroDocumento = Format(MovCupomFiscal.NumeroCupom, "#######0") & Format(MovCupomFiscal.Ordem, "00")
            
            If UCase(g_nome_empresa) Like "*POSTO T13*" Or UCase(g_nome_empresa) Like "*MARQUES*" Or UCase(g_nome_empresa) Like "*CLASSE A*" Then
               If Val(txt_numero_nota_abastecimento.Text) > 0 Then
                    MovCaixaPista.NumeroDocumento = txt_numero_nota_abastecimento.Text
               End If
            End If

        ElseIf pTipoLancamentoPadrao = "CartaoCredito" Then
            MovCaixaPista.Valor = lValorTotalUltimoCupom
            MovCaixaPista.CodigoHistorico = IntegracaoCaixa.HistoricoPadrao
            MovCaixaPista.DadosInterno = "CAR" & Format(CartaoCredito.Codigo, "00") & "|@|" & lNumeroLancamentoCartao & "|@|"
            xComplemento = "P/ " & CDate(MovCupomFiscal.Data + CartaoCredito.DiasPrazo) & " TM:" & MovCaixaPista.TipoMovimento & " P:" & MovCupomFiscal.Periodo
            MovCaixaPista.CodigoLancamentoPadrao = 2
            MovCaixaPista.NumeroDocumento = Format(MovCupomFiscal.NumeroCupom, "#######0") & Format(MovCupomFiscal.Ordem, "00")
        ElseIf pTipoLancamentoPadrao = "VENDA DE LUBRIFICANTES" Then
            MovCaixaPista.Valor = xValor
            MovCaixaPista.CodigoHistorico = IntegracaoCaixa.HistoricoPadrao
            MovCaixaPista.DadosInterno = "LUBRI" & "|@|" & Val(cboTipoSubEstoque.Text) & "|@|"
            xComplemento = "LUBRIFICANTES Per:" & MovCupomFiscal.Periodo & " Ilha:" & lIlha & " S.Est:" & Val(cboTipoSubEstoque.Text) & " T.Mov:" & lTipoMovimento
            MovCaixaPista.CodigoLancamentoPadrao = 1
            MovCaixaPista.NumeroDocumento = ""
        ElseIf pTipoLancamentoPadrao = "DESCONTO AUTORIZADO" Then
            MovCaixaPista.Valor = pValor
            MovCaixaPista.CodigoHistorico = IntegracaoCaixa.HistoricoPadrao
            MovCaixaPista.DadosInterno = "DESC.ABAST|@|" & pComplementoDadosInterno
            MovCaixaPista.CodigoLancamentoPadrao = 23
            MovCaixaPista.NumeroDocumento = ""
            xComplemento = pComplementoCaixa
        ElseIf pTipoLancamentoPadrao = "DINHEIRO" Then
            MovCaixaPista.Valor = pValor
            MovCaixaPista.CodigoHistorico = IntegracaoCaixa.HistoricoPadrao
            MovCaixaPista.DadosInterno = "CX.PISTA DIV." & "|@|"
            xComplemento = pComplementoCaixa
            MovCaixaPista.CodigoLancamentoPadrao = 8
            MovCaixaPista.NumeroDocumento = ""
        ElseIf pTipoLancamentoPadrao = "CHEQUE A VISTA" Then
            MovCaixaPista.Valor = pValor
            MovCaixaPista.CodigoHistorico = IntegracaoCaixa.HistoricoPadrao
            MovCaixaPista.DadosInterno = "CX.PISTA DIV." & "|@|"
            xComplemento = pComplementoCaixa
            MovCaixaPista.CodigoLancamentoPadrao = 13
            MovCaixaPista.NumeroDocumento = txt_numero_cheque.Text
        ElseIf pTipoLancamentoPadrao = "DescontoCartaoFidelidade" Then
            MovCaixaPista.Valor = xValorDesconto
            xComplemento = pComplementoCaixa
            MovCaixaPista.CodigoHistorico = IntegracaoCaixa.HistoricoPadrao
            MovCaixaPista.DadosInterno = "C.FID|@|"
            MovCaixaPista.CodigoLancamentoPadrao = 31
            MovCaixaPista.NumeroDocumento = Format(MovCupomFiscal.NumeroCupom, "#######0") & Format(MovCupomFiscal.Ordem, "00")
        ElseIf pTipoLancamentoPadrao = "DescontoCartaoCorreios" Then
            MovCaixaPista.Valor = xValorDesconto
            xComplemento = pComplementoCaixa
            MovCaixaPista.CodigoHistorico = IntegracaoCaixa.HistoricoPadrao
            MovCaixaPista.DadosInterno = "C.CORREIOS|@|"
            MovCaixaPista.CodigoLancamentoPadrao = 32
            MovCaixaPista.NumeroDocumento = Format(MovCupomFiscal.NumeroCupom, "#######0") & Format(MovCupomFiscal.Ordem, "00")
        ElseIf pTipoLancamentoPadrao = "DescontoCartaoEspecial" Then
            MovCaixaPista.Valor = xValorDesconto
            xComplemento = pComplementoCaixa
            MovCaixaPista.CodigoHistorico = IntegracaoCaixa.HistoricoPadrao
            MovCaixaPista.DadosInterno = "C.ESPECIAL|@|"
            MovCaixaPista.CodigoLancamentoPadrao = 34
            MovCaixaPista.NumeroDocumento = Format(MovCupomFiscal.NumeroCupom, "#######0") & Format(MovCupomFiscal.Ordem, "00")
        End If
        
        MovCaixaPista.Empresa = g_empresa
        MovCaixaPista.Data = pData
        MovCaixaPista.NumeroMovimento = 1
        MovCaixaPista.Complemento = Mid(xComplemento, 1, 50)
        MovCaixaPista.NumeroContaDebito = xContaDebito
        MovCaixaPista.NumeroContaCredito = xContaCredito
        MovCaixaPista.CodigoUsuario = g_usuario
        MovCaixaPista.TipoMovimento = lTipoMovimento
        MovCaixaPista.Periodo = pPeriodo
        MovCaixaPista.NumeroIlha = lIlha
        If pDesconto And (pTipoLancamentoPadrao <> "DescontoCartaoFidelidade" And pTipoLancamentoPadrao <> "DescontoCartaoCorreios" And pTipoLancamentoPadrao <> "DescontoCartaoEspecial") Then
            MovCaixaPista.DadosInterno = ""
        End If
        MovCaixaPista.DataDigitacao = Format(Now, "dd/mm/yyyy")
        MovCaixaPista.HoraDigitacao = Format(Now, "HH:mm:ss")
        MovCaixaPista.DataAlteracao = "00:00:00"
        MovCaixaPista.HoraAlteracao = "00:00:00"
        If MovCaixaPista.Incluir Then
            'lNumeroMovimentoCaixa = MovCaixaPista.NumeroMovimento
            IncluiMovimentoCaixa = True
        End If
    Else
        MsgBox "Não será possível integrar com o caixa!", vbInformation + vbCritical, "Erro de Integridade"
    End If
End Function
Private Function EstornaVendaProdutoNFCe(ByVal pNumeroNFCe As Long, ByVal pDataEmissao As Date, ByVal pPeriodo As Integer, ByVal pCodigoProduto As Integer, ByVal pCodigoFuncionario As Integer, ByVal pQuantidade As Currency, ByVal pValorTotal As Currency, ByVal pTipoSubEstoque As Integer) As Boolean

On Error GoTo trata_erro
    
    EstornaVendaProdutoNFCe = False
    
    If Not IntegracaoCaixa.LocalizarNome(g_empresa, "VENDA DE LUBRIFICANTES") Then
        MsgBox "Não será possível integrar com o caixa!", vbCritical, "Erro de Integridade!"
        Call GravaAuditoria(1, Me.name, 25, "Não será integrado no caixa o extorno de produto.")
        Exit Function
    End If
        
    If ExcluiMovimentoCaixaNFCe("VENDA DE LUBRIFICANTES", pDataEmissao, pPeriodo, pValorTotal, pTipoSubEstoque) Then
        If MovimentoLubrificante.LocalizarCodigo(g_empresa, pDataEmissao, pPeriodo, lIlha, lTipoMovimento, pTipoSubEstoque, pCodigoProduto, pCodigoFuncionario) Then
            MovimentoLubrificante.Quantidade = MovimentoLubrificante.Quantidade - pQuantidade
            MovimentoLubrificante.ValorTotal = MovimentoLubrificante.ValorTotal - pValorTotal
            If MovimentoLubrificante.Quantidade = 0 Then
                If MovimentoLubrificante.Excluir(g_empresa, pDataEmissao, pPeriodo, lIlha, lTipoMovimento, pTipoSubEstoque, pCodigoProduto, pCodigoFuncionario) Then
                    EstornaVendaProdutoNFCe = True
                Else
                    Call GravaAuditoria(1, Me.name, 25, "Não excluiu venda de produto:" & pCodigoProduto)
                    Call CriaLogCupom(Time & "EstornaVendaProdutoNFCe: Não excluiu venda de produto:" & pCodigoProduto & " Data:" & pDataEmissao & " Per:" & pPeriodo & " Ilha:" & lIlha & " Tipo Mov:" & lTipoMovimento & " SubEst:" & pTipoSubEstoque & " Prod:" & pCodigoProduto & " Operador:" & pCodigoFuncionario)
                    MsgBox "Não foi possível excluir venda de produtos.", vbCritical, "Erro de Integridade!"
                End If
            Else
                If MovimentoLubrificante.Alterar(g_empresa, pDataEmissao, pPeriodo, lIlha, lTipoMovimento, pTipoSubEstoque, pCodigoProduto, pCodigoFuncionario) Then
                    EstornaVendaProdutoNFCe = True
                Else
                    Call GravaAuditoria(1, Me.name, 25, "Não alterou venda de produto:" & pCodigoProduto)
                    Call CriaLogCupom(Time & "EstornaVendaProdutoNFCe: Não alterou venda de produto:" & pCodigoProduto & " Data:" & pDataEmissao & " Per:" & pPeriodo & " Ilha:" & lIlha & " Tipo Mov:" & lTipoMovimento & " SubEst:" & pTipoSubEstoque & " Prod:" & pCodigoProduto & " Operador:" & pCodigoFuncionario)
                    MsgBox "Não foi possível alterar venda de produtos.", vbCritical, "Erro de Integridade!"
                End If
            End If
        Else
            Call GravaAuditoria(1, Me.name, 25, "Não localizou venda de produto:" & pCodigoProduto)
            Call CriaLogCupom(Time & "SubtraiVendaProduto: Não localizou venda de produto:" & pCodigoProduto & " Data:" & pDataEmissao & " Per:" & pPeriodo & " Ilha:" & lIlha & " Tipo Mov:" & lTipoMovimento & " SubEst:" & pTipoSubEstoque & " Prod:" & pCodigoProduto & " Operador:" & pCodigoFuncionario)
            MsgBox "Não foi possível localizar venda de produtos.", vbCritical, "Erro de Integridade!"
        End If
    Else
        Call GravaAuditoria(1, Me.name, 25, "Não foi possível estornar venda de produto no caixa.")
        MsgBox "Não foi possível estornar no caixa!", vbCritical, "Erro de Integridade!"
    End If
    Exit Function

trata_erro:
    Call CriaLogCupom("Erro EstornaVendaProdutoNFCe: Erro=" & Err.Number & " - " & Err.Description)
    Call GravaAuditoria(1, Me.name, 25, "EstornaVendaProdutoNFCe: Erro inesperado...")
End Function

Private Sub IncluiNotaAbastecimento()
On Error GoTo trata_erro
    
    Dim rsMovCupomFiscal As New ADODB.Recordset
    
    If Cliente.GeraNotaAbastecimento Then
        lSQL = ""
        lSQL = lSQL & "SELECT * FROM Movimento_Cupom_Fiscal_Item"
        lSQL = lSQL & " WHERE Empresa = " & g_empresa
        lSQL = lSQL & "   AND [Numero do Cupom] = " & lNumeroCupom
        lSQL = lSQL & "   AND Data = " & preparaData(lData)
        lSQL = lSQL & " ORDER BY Ordem"
        Set rsMovCupomFiscal = Conectar.RsConexao(lSQL)
        If rsMovCupomFiscal.RecordCount > 0 Then
            Do Until rsMovCupomFiscal.EOF
                If Not IntegracaoCaixa.LocalizarNome(g_empresa, "NOTA ABASTECIMENTO") Then
                    Call CriaLogCupom("Erro AtualizaTabelaNotaAbastecimento:Integração de caixa inexistente. Cliente=" & rsMovCupomFiscal![Codigo do Cliente])
                    MsgBox "Não será possível integrar com o caixa!", vbInformation, "Erro de Integridade"
                Else
                    If Not MovCupomFiscal.LocalizarCodigo(g_empresa, lNumeroPDV, rsMovCupomFiscal![Numero do Cupom], rsMovCupomFiscal!Data, rsMovCupomFiscal!Ordem) Then
                        MsgBox "Não foi possível localizar movimento de cupom fiscal!", vbInformation, "Erro de Integridade"
                    End If
                    If IncluiMovimentoCaixa(MovCupomFiscal.Data, MovCupomFiscal.Periodo, False, "NotaAbastecimento", 0, "", "") Then
                        MovNotaAbastecimento.Empresa = g_empresa
                        MovNotaAbastecimento.CodigoCliente = l_codigo_cliente 'rsMovCupomFiscal![Codigo do Cliente]
                        MovNotaAbastecimento.DataAbastecimento = rsMovCupomFiscal!Data
                        MovNotaAbastecimento.NumeroNota = Format(rsMovCupomFiscal![Numero do Cupom], "00000000") & Format(rsMovCupomFiscal!Ordem, "00")
                        'aqui numero da nota
                        If UCase(g_nome_empresa) Like "*POSTO T13*" Or UCase(g_nome_empresa) Like "*MARQUES*" Or UCase(g_nome_empresa) Like "*CLASSE A*" Then
                            If Val(txt_numero_nota_abastecimento.Text) > 0 Then
                                MovNotaAbastecimento.NumeroNota = txt_numero_nota_abastecimento.Text
                            End If
                        End If

                        'aqui numero da nota
                        'If UCase(g_nome_empresa) Like "*MARQUES*" Or UCase(g_nome_empresa) Like "*RATINHO*" Then
                        '    MovNotaAbastecimento.NumeroNota = txt_numero_nota_abastecimento.Text
                        'End If
                        MovNotaAbastecimento.Ordem = rsMovCupomFiscal!Ordem
                        MovNotaAbastecimento.CodigoProduto2 = rsMovCupomFiscal![Codigo do Produto]
                        MovNotaAbastecimento.Periodo = rsMovCupomFiscal!Periodo
                        MovNotaAbastecimento.Quantidade = rsMovCupomFiscal!Quantidade
                        
                        'MovNotaAbastecimento.ValorUnitario = rsMovCupomFiscal![Valor Unitario] - rsMovCupomFiscal![Valor do Acrescimo] + DescontoPersonalizado(l_codigo_cliente, Cliente.CodigoGrupoCliente, rsMovCupomFiscal![Codigo do Produto], rsMovCupomFiscal![Valor Unitario])
                        
                        'A linha abaixo corrigia o problema do valor unitário negativo
                        'que acontecia na linha acima comentada
                        'MovNotaAbastecimento.ValorUnitario = rsMovCupomFiscal![Valor Unitario] - DescontoPersonalizado(l_codigo_cliente, Cliente.CodigoGrupoCliente, rsMovCupomFiscal![Codigo do Produto], rsMovCupomFiscal![Valor Unitario])
                        
                        'O Valor Unitario, teoricamente deve ser o valor unitário vindo cupom fiscal
                        'ou seja, o valor unitário do produto - desconto ou + acrescimo
                        MovNotaAbastecimento.ValorUnitario = rsMovCupomFiscal![Valor Unitario]
                        MovNotaAbastecimento.ValorTotal = lTotalItem(rsMovCupomFiscal!Ordem) 'lValorTotalSemDesconto 'rsMovCupomFiscal![Valor Total]
                        Call CriaLogCupom("Monitora-01 Total Nota. (Origem Variável lTotalItem index(" & rsMovCupomFiscal!Ordem & ") ) N.Cupom:" & rsMovCupomFiscal![Numero do Cupom] & " - Quantidade:" & rsMovCupomFiscal!Quantidade & " - Valor Total:" & rsMovCupomFiscal![Valor Total])
                        If MovNotaAbastecimento.ValorTotal = 0 Then
                            MovNotaAbastecimento.ValorTotal = rsMovCupomFiscal![Valor Total]
                            Call CriaLogCupom("Monitora-02 Total Nota. (MovNotaAbastecimento.ValorTotal = 0. Origem Cupom) N.Cupom:" & rsMovCupomFiscal![Numero do Cupom] & " - Quantidade:" & rsMovCupomFiscal!Quantidade & " - Valor Total:" & rsMovCupomFiscal![Valor Total])
                        End If
                        If MovNotaAbastecimento.ValorTotal = 0 Then
                            MovNotaAbastecimento.ValorTotal = MovNotaAbastecimento.ValorUnitario * MovNotaAbastecimento.Quantidade
                            Call CriaLogCupom("Monitora-03 Total Nota. (MovNotaAbastecimento.ValorTotal = 0. MovNotaAbastecimento: Calculo Qtd * Vlr Unitário) N.Cupom:" & rsMovCupomFiscal![Numero do Cupom] & " - MovNotaAbastecimento.Quantidade:" & MovNotaAbastecimento.Quantidade & " - MovNotaAbastecimento.ValorUnitario:" & MovNotaAbastecimento.ValorUnitario)
                        End If
                        MovNotaAbastecimento.CodigoConveniado = 0
                        MovNotaAbastecimento.TipoMovimento = lTipoMovimento '2-Pista, 1-Conveniencia rsMovCupomFiscal![Tipo do Movimento] - 1
                        MovNotaAbastecimento.PlacaLetra = lPlacaLetra
                        MovNotaAbastecimento.PlacaNumero = lPlacaNumero
                        If Not Trim(txt_placa.Text) = Empty Then
                            MovNotaAbastecimento.PlacaLetra = Mid(txt_placa.Text, 1, 3)
                            If Len(Trim(txt_placa.Text)) = 7 Then
                                MovNotaAbastecimento.PlacaNumero = Mid(txt_placa.Text, 4, 4)
                            Else
                                MovNotaAbastecimento.PlacaNumero = Mid(txt_placa.Text, 5, 4)
                            End If
                        End If
                        MovNotaAbastecimento.Historico = "E.C.F."
                        MovNotaAbastecimento.NumeroCupom = rsMovCupomFiscal![Numero do Cupom]
                        MovNotaAbastecimento.ValorDescontoUnitario = DescontoPersonalizado(l_codigo_cliente, Cliente.CodigoGrupoCliente, rsMovCupomFiscal![Codigo do Produto], rsMovCupomFiscal![Valor Unitario])
                        'a linha abaixo talvez resolva o problema no posto granadinha
                        'pois teoricamente o valor deve ser sem desconto.
                        If MovNotaAbastecimento.ValorDescontoUnitario <> 0 Then
                            MovNotaAbastecimento.ValorUnitario = rsMovCupomFiscal![Valor Unitario] + MovNotaAbastecimento.ValorDescontoUnitario
                            MovNotaAbastecimento.ValorTotal = MovNotaAbastecimento.ValorUnitario * MovNotaAbastecimento.Quantidade
                            Call CriaLogCupom("Monitora-04 Total Nota. (MovNotaAbastecimento.ValorDescontoUnitario <> 0. MovNotaAbastecimento: Calculo Qtd * Vlr Unitário) N.Cupom:" & rsMovCupomFiscal![Numero do Cupom] & " - MovNotaAbastecimento.Quantidade:" & MovNotaAbastecimento.Quantidade & " - MovNotaAbastecimento.ValorUnitario:" & MovNotaAbastecimento.ValorUnitario)
                        End If
                        MovNotaAbastecimento.NumeroMovimentoCaixa = MovCaixaPista.NumeroMovimento
                        MovNotaAbastecimento.BaixadoPelaDuplicata = False
                        MovNotaAbastecimento.NumeroIlha = lIlha
                        MovNotaAbastecimento.Origem = "CF"
                        MovNotaAbastecimento.DataConferencia = "00:00:00"
                        MovNotaAbastecimento.KM = lKMVeiculo
                        If Not Trim(txt_kilometragem.Text) = Empty Then
                            MovNotaAbastecimento.KM = Val(txt_kilometragem.Text)
                        End If
                        
                        If MovNotaAbastecimento.Incluir Then
                            If MovNotaAbastecimento.ValorDescontoUnitario <> 0 Then
                                If Not IncluiMovimentoCaixa(MovCupomFiscal.Data, MovCupomFiscal.Periodo, True, "NotaAbastecimento", 0, "", "") Then
                                    Call CriaLogCupom("Erro AtualizaTabelaNotaAbastecimento:Desconto/Acréscimo não integrada no caixa. Cliente=" & MovCupomFiscal.CodigoCliente)
                                    MsgBox "Não foi possível integrar Desconto/Acréscimo no caixa!", vbInformation, "Erro de Integridade!"
                                End If
                            End If
                        Else
                            Call CriaLogCupom("Erro IncluiNotaAbasteciment:Não gravada. Cliente=" & l_codigo_cliente)
                            MsgBox "Não foi possível incluir Nota de Abastecimento", vbInformation, "Erro de Integridade!"
                        End If
                    Else
                        Call CriaLogCupom("Erro IncluiNotaAbastecimento:Não integrada no caixa. Cliente=" & MovCupomFiscal.CodigoCliente)
                        MsgBox "Não foi possível integrar no caixa!", vbInformation, "Erro de Integridade!"
                    End If
                End If
                rsMovCupomFiscal.MoveNext
            Loop
        End If
        rsMovCupomFiscal.Close
    End If
    Exit Sub

trata_erro:
    Call CriaLogCupom("Erro IncluiNotaAbasteciment:Desconhecido. Cliente=" & rsMovCupomFiscal![Codigo do Cliente])
    Call CriaLogCupom("Erro IncluiNotaAbasteciment: Erro=" & Err.Number & " - " & Err.Description)
    MsgBox Err.Number & " - " & Err.Description
End Sub
Private Sub IncluiSaidaBomba(x_produto As Long, x_quantidade As Currency)
    Dim xTipoCombustivel As String
    xTipoCombustivel = Bomba.LocalizarCodigoProduto(g_empresa, x_produto)
    If Combustivel.LocalizarCodigo(g_empresa, xTipoCombustivel) Then
        Combustivel.QuantidadeEmEstoque = Combustivel.QuantidadeEmEstoque - x_quantidade
        If Not Combustivel.Alterar(g_empresa, xTipoCombustivel) Then
            MsgBox "Não foi possível alterar registro de combustível!", vbInformation, "Erro de Integridade!"
        End If
    End If
End Sub
Private Sub IncluiSaidaProduto(pCodigoProduto As Long, pQuantidade As Currency)
    If Estoque.LocalizarCodigo(g_empresa, pCodigoProduto) Then
        Estoque.Quantidade = Estoque.Quantidade - pQuantidade
        If Not Estoque.Alterar(g_empresa, pCodigoProduto) Then
            MsgBox "Não foi possível alterar o estoque!", vbInformation, "Erro de Integridade!"
        End If
    Else
        MsgBox "Estoque não cadastrado.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub IniciaContadorAguarde(ByVal pNumeroInicial As Integer, Optional pEspacoContador As Integer = 1000)
    lblContadorAguarde.Visible = True
    lblContadorAguarde.Width = Me.Width - 600
    lblContadorAguarde.Top = lblMensagemAguarde.Top + pEspacoContador
    lblContadorAguarde.Left = 200
    lContadorAguarde = pNumeroInicial
    lblContadorAguarde.Caption = lContadorAguarde
    DoEvents
    TimerAguarde.Enabled = True
    TimerAguarde.Interval = 1000
End Sub
Private Sub IniciaProcessamentoNFCe()
    Dim i As Integer
    Dim xString As String
    Dim xImprimeTef As Boolean
    Dim xDadosProdutos As String
    Dim xObservacao2 As String
    Dim xLinhaImpostos As String
    Dim xTextoParaComprovante As String
    Dim xFechamentoIniciado As Boolean
    Dim xTipoServico As String
    
    xTipoServico = "NFCe 3.10"
    
    If ConfiguracaoDiversa.LocalizarCodigo(1, "VERSAO NFCE") Then
        xTipoServico = "NFCe" & " " & ConfiguracaoDiversa.Texto
    End If

    
    If MovSolicitacaoFuncaoNFe.ExisteSolicitacaoParaOperacao(Now, lNumeroCupom, lSerieNFCe, xTipoServico) Then
        Call CriaLogCupom(Time & " - [IniciaProcessamentoNFCe] - O sistema encontrou uma solicitação para o lNumeroCupom=" & lNumeroCupom & " POR ISSO IMPEDIRÁ ESSE PROCESSAMENTO")
        Call CriaLogMonitoraBug("IniciaProcessamentoNFCe", "Movimento_NfceAuto", "Nº NFCe já existente")
        cmdIniciaProcessoNFCe.Enabled = True
        Exit Sub
    End If
    
    i = 0
    xImprimeTef = False
    xLinhaImpostos = ""
    xFechamentoIniciado = False
            
    DefineImpressoraTermicaComoPadrao
    
    If ValidaCampos2 Then
    '25/06/14^
        frmDados.Enabled = True
        cmdIniciaProcessoFinalizacaoNFCe.Visible = True
        lValorTotalUltimoCupom = fValidaValor(Me.lbl_valor_compra.Caption)
        Call GravaAuditoria(1, Me.name, 23, "ECF fechado em:" & Me.cbo_forma_pagamento.Text & " Vlr.Recebido:" & txt_valor_recebido.Text)
        
        If lExigeNCM = True Then
            xLinhaImpostos = CalculaImpostos(lNumeroCupom, lData)
        End If
        If lTEF Then
            If cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 4 Or cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 6 Or cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 7 Or cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 8 Or cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 9 Or cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 10 Or cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 11 Or cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 12 Or cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 13 Or cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 14 Or cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 15 Or cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 17 Then
                Call CriaLogECF(Date & " " & Time & " TEF: N.Cupom=" & lNumeroCupom & " - Valor=" & txt_valor_recebido.Text & " - Forma Pg.=" & cbo_forma_pagamento.Text)
                gNumeroControleSolicitacao = Configuracao.ProximaSolicitacaoTEF(g_empresa)
                'Call TestaConexao(1, "IniciaProcessoNFCe")
                lRespostaTEF = False
                'Set CerradoTef = Nothing
                'Call TestaConexao(2, "IniciaProcessoNFCe")
                Set CerradoTef = New CerradoComponenteTef
                'Call TestaConexao(3, "IniciaProcessoNFCe")
                If lNumeroCupom <> lNumeroUltimoCupom Then
                    MsgBox "ERRO DO NUMERO DO CUPOM:" & lNumeroUltimoCupom & " <> " & lNumeroCupom
                End If
                If txt_observacao_2.Text = "" And txt_placa.Text <> "" Then
                '25/06/14^
                    xString = "PLACA.:             KILOMETRAGEM..:             "
                    Mid(xString, 9, 8) = txt_placa.Text
                    Mid(xString, 37, 12) = txt_kilometragem.Text
                    txt_observacao_2.Text = xString
                End If
                xObservacao2 = xLinhaImpostos & txt_observacao_2.Text
                '25/06/14^
                
                'Prepara Texto para sair no comprovante de venda
                'aqui
                xTextoParaComprovante = "Funcionario:" + Format(l_codigo_funcionario, "000") + " " + Mid(l_nome_funcionario, 1, 15)
                If Len(xTextoParaComprovante) < 48 Then
                    Do While Len(xTextoParaComprovante) <= 48
                        xTextoParaComprovante = xTextoParaComprovante & " "
                    Loop
                End If
                xTextoParaComprovante = String(48, "-") & xTextoParaComprovante & String(48, "-")
                '
                'Teste cartao: ao chegar aqui pular para o ponto B
                'e mudar o valor da variavel lRespostaTEF para true
                If lValorDescontoConcedido > 0 Then
                    xFechamentoIniciado = True
                End If
                If cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 4 Then
                    lRespostaTEF = CerradoTef.SolicitacaoTEF("NFCe", gNumeroControleSolicitacao, lNumeroCupom, txt_valor_recebido.Text, txt_valor_desconto.Text, gQtdViasTEF, "Outras", False, txt_cpf.Text, txt_nome_cliente.Text, txt_observacao.Text, xObservacao2, lLinhasEntreCV, xTextoParaComprovante, xFechamentoIniciado, l_codigo_funcionario, l_nome_funcionario)
                ElseIf cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 6 Then
                    lRespostaTEF = CerradoTef.SolicitacaoTEF("NFCe", gNumeroControleSolicitacao, lNumeroCupom, txt_valor_recebido.Text, txt_valor_desconto.Text, gQtdViasTEF, "TecBan", False, txt_cpf.Text, txt_nome_cliente.Text, txt_observacao.Text, xObservacao2, lLinhasEntreCV, xTextoParaComprovante, xFechamentoIniciado, l_codigo_funcionario, l_nome_funcionario)
                ElseIf cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 7 Then
                    lRespostaTEF = CerradoTef.SolicitacaoTEF("NFCe", gNumeroControleSolicitacao, lNumeroCupom, txt_valor_recebido.Text, txt_valor_desconto.Text, gQtdViasTEF, "TecBan", True, txt_cpf.Text, txt_nome_cliente.Text, txt_observacao.Text, xObservacao2, lLinhasEntreCV, xTextoParaComprovante, xFechamentoIniciado, l_codigo_funcionario, l_nome_funcionario)
                ElseIf cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 8 Then
                    Call MontaDadosTCS(lNumeroCupom, lData)
                    lRespostaTEF = CerradoTef.SolicitacaoTefTCS("NFCe", gNumeroControleSolicitacao, lNumeroCupom, txt_valor_recebido.Text, txt_valor_desconto.Text, gQtdViasTEF, lDadosTCS, lLegislacaoPermiteIssEcf, lCodigoTcsEcf, lContadorNaoFiscal, txt_cpf.Text, txt_nome_cliente.Text, txt_observacao.Text, xObservacao2, lLinhasEntreCV, xTextoParaComprovante, l_codigo_funcionario, l_nome_funcionario)
                ElseIf cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 9 Then
                    lRespostaTEF = CerradoTef.SolicitacaoTEF("NFCe", gNumeroControleSolicitacao, lNumeroCupom, txt_valor_recebido.Text, txt_valor_desconto.Text, gQtdViasTEF, "SMARTEF", False, txt_cpf.Text, txt_nome_cliente.Text, txt_observacao.Text, xObservacao2, lLinhasEntreCV, xTextoParaComprovante, xFechamentoIniciado, l_codigo_funcionario, l_nome_funcionario)
                ElseIf cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 10 Then
                    lRespostaTEF = CerradoTef.SolicitacaoTEF("NFCe", gNumeroControleSolicitacao, lNumeroCupom, txt_valor_recebido.Text, txt_valor_desconto.Text, gQtdViasTEF, "SUPERTEF", False, txt_cpf.Text, txt_nome_cliente.Text, txt_observacao.Text, xObservacao2, lLinhasEntreCV, xTextoParaComprovante, xFechamentoIniciado, l_codigo_funcionario, l_nome_funcionario)
                ElseIf cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 11 Then
                    lRespostaTEF = CerradoTef.SolicitacaoTEF("NFCe", gNumeroControleSolicitacao, lNumeroCupom, txt_valor_recebido.Text, txt_valor_desconto.Text, gQtdViasTEF, "HIPERTEF", False, txt_cpf.Text, txt_nome_cliente.Text, txt_observacao.Text, xObservacao2, lLinhasEntreCV, xTextoParaComprovante, xFechamentoIniciado, l_codigo_funcionario, l_nome_funcionario)
                ElseIf cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 12 Then
                    lRespostaTEF = CerradoTef.SolicitacaoTEF("NFCe", gNumeroControleSolicitacao, lNumeroCupom, txt_valor_recebido.Text, txt_valor_desconto.Text, gQtdViasTEF, "PAGCARD", False, txt_cpf.Text, txt_nome_cliente.Text, txt_observacao.Text, xObservacao2, lLinhasEntreCV, xTextoParaComprovante, xFechamentoIniciado, l_codigo_funcionario, l_nome_funcionario)
                ElseIf cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 13 Then
                    lRespostaTEF = CerradoTef.SolicitacaoTEF("NFCe", gNumeroControleSolicitacao, lNumeroCupom, txt_valor_recebido.Text, txt_valor_desconto.Text, gQtdViasTEF, "CHEQUEREDECARD", True, txt_cpf.Text, txt_nome_cliente.Text, txt_observacao.Text, xObservacao2, lLinhasEntreCV, xTextoParaComprovante, xFechamentoIniciado, l_codigo_funcionario, l_nome_funcionario)
                ElseIf cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 14 Then
                    lRespostaTEF = CerradoTef.SolicitacaoTEF("NFCe", gNumeroControleSolicitacao, lNumeroCupom, txt_valor_recebido.Text, txt_valor_desconto.Text, gQtdViasTEF, "TEFNEUS", False, txt_cpf.Text, txt_nome_cliente.Text, txt_observacao.Text, xObservacao2, lLinhasEntreCV, xTextoParaComprovante, xFechamentoIniciado, l_codigo_funcionario, l_nome_funcionario)
                ElseIf cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 15 Then
                    lRespostaTEF = CerradoTef.SolicitacaoTEF("NFCe", gNumeroControleSolicitacao, lNumeroCupom, txt_valor_recebido.Text, txt_valor_desconto.Text, gQtdViasTEF, "GODCARD", False, txt_cpf.Text, txt_nome_cliente.Text, txt_observacao.Text, xObservacao2, lLinhasEntreCV, xTextoParaComprovante, xFechamentoIniciado, l_codigo_funcionario, l_nome_funcionario)
                ElseIf cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 17 Then
                    xDadosProdutos = xObservacao2 & vbCrLf & PreparaDadosProdutos
                    lRespostaTEF = CerradoTef.SolicitacaoTEF("NFCe", gNumeroControleSolicitacao, lNumeroCupom, txt_valor_recebido.Text, txt_valor_desconto.Text, gQtdViasTEF, "TEFCERRADO", False, txt_cpf.Text, txt_nome_cliente.Text, xObservacao2 & txt_observacao.Text, xDadosProdutos, lLinhasEntreCV, xTextoParaComprovante, xFechamentoIniciado, l_codigo_funcionario, l_nome_funcionario)
                End If
                'Call TestaConexao(4, "IniciaProcessoNFCe")
                'PONTO B
                Call CriaLogECF(Date & " " & Time & " TEF: N.Cupom=" & lNumeroCupom & " - Retorno lRespostaTEF=" & lRespostaTEF)
'                If txt_observacao.Text = "Funcionario:" + Format(l_codigo_funcionario, "000") + " " + Mid(l_nome_funcionario, 1, 15) Then
'                    txt_observacao.Text = ""
'                ElseIf txt_observacao_2.Text = "Funcionario:" + Format(l_codigo_funcionario, "000") + " " + Mid(l_nome_funcionario, 1, 15) Then
'                    txt_observacao_2.Text = ""
'                End If
                
'                If lImpQuick Then 'ALEX - NFCE
'                    If lRespostaTEF Then
'                        If EcfQuickLeRegistrador("EstadoFiscal", "Long", 5) = "" Then
'                            lRespostaTEF = False
'                        End If
'                        If EcfQuickLeRegistrador("EstadoFiscal", "Long", 5) = "2" Then
'                            lRespostaTEF = False
'                        End If
'                    End If
'                End If
                
                
                'teste para fechar gerencial caso esteja aberto
'                If lImpQuick Then 'ALEX - NFCE
'                    If EcfQuickLeRegistrador("EstadoFiscal", "Long", 5) = "32" Then
'                        Call EcfQuickEncerraDocumento(0, "Gerencial")
'                    End If
'                    If EcfQuickLeRegistrador("EstadoFiscal", "Long", 5) = "64" Then
'                        Call EcfQuickEncerraDocumento(0, "Gerencial")
'                    End If
'                End If
'
                
                Set CerradoTef = Nothing
                'Call TestaConexao(5, "IniciaProcessoNFCe")
                If lRespostaTEF = True Then
                    xImprimeTef = True
                    'Call TestaConexao(6, "IniciaProcessoNFCe")
                    If IntegraCartaoCreditoNoCaixa Then
                        'Call TestaConexao(7, "IniciaProcessoNFCe")
                        DefineCartaoTefParaNFCe
                        'Call TestaConexao(8, "IniciaProcessoNFCe")
                        AtualizaTabelaCartaoCredito
                        ' Desconto Cartao Correios
                        If lIntegraDescontoCartaoCorreios = True And lValorDescontoConcedido > 0 Then
                            If Not IncluiMovimentoCaixa(lDataCupom, lPeriodo, True, "DescontoCartaoCorreios", lValorDescontoConcedido, "", "Cartão Desconto Correios") Then
                                Call CriaLogCupom("Erro cmd_DescontoCorreio_Click:Desconto Cartao Correios não integrada no caixa.")
                                MsgBox "Não foi possível integrar Desconto Cartão Correios no caixa!", vbInformation, "Erro de Integridade!"
                            End If
                        End If
                    Else
                        'Call TestaConexao(9, "IniciaProcessoNFCe")
                        DefineCartaoTefParaNFCe
                    End If
                Else
                    'Teste para rastrear bug que imprime o Comprovante
                    ' e pede para o usuario passar novamente o cartao
                    If g_nome_empresa Like "*RATINHO*" Then
                        Dim xArqTxt As New FileSystemObject
                        Dim xNomeArquivo As String
                        Dim xNomeArquivoCopia As String
                        xNomeArquivo = "c:\vb5\sgp\data\teste.txt"
                        xNomeArquivoCopia = "TTF_" & Format(Date, "dd") & "_" & Format(Date, "MM") & "_" & Format(Date, "yyyy") & "__" & Format(Now, "HH:mm:ss") & ".LOG"
                        Mid(xNomeArquivoCopia, 19, 1) = "_"
                        Mid(xNomeArquivoCopia, 22, 1) = "_"
                        xNomeArquivoCopia = "c:\vb5\sgp\data\" & xNomeArquivoCopia
                        If xArqTxt.FileExists(xNomeArquivo) Then
                            Call xArqTxt.CopyFile(xNomeArquivo, xNomeArquivoCopia, True)
                        End If
                    End If
                    'fim do teste do bug
                    MsgBox "Selecione outra forma de pagamento!", vbInformation, "Forma de Pagamento Temporariamente Não Aceita!"
                    cbo_forma_pagamento.SetFocus
                    cmdIniciaProcessoNFCe.Enabled = True
                    Exit Sub
                End If
            End If
        End If
        
        'Call TestaConexao(10, "IniciaProcessoNFCe")
        MovCupomFiscal.FormaPagamento = cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex)
        MovCupomFiscal.ValorRecebido = fValidaValor2(txt_valor_recebido.Text)
        'MovCupomFiscal.NumeroCheque = txt_numero_cheque.Text
        'MovCupomFiscal.Telefone = fDesmascaraTelefone(txt_telefone.Text)
        MovCupomFiscal.operador = l_codigo_funcionario
        MovCupomFiscal.CodigoCliente = l_codigo_cliente
        
        'ALEX-DOCELETRONICO
        MovDocEletronicoCabecalho.FormaPagamento = cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex)
        MovDocEletronicoCabecalho.ValorTotal = fValidaValor2(lbl_valor_compra.Caption)
        MovDocEletronicoCabecalho.ValorDesconto = fValidaValor2(txt_valor_desconto.Text)
        MovDocEletronicoCabecalho.ValorProdutos = fValidaValor2(lbl_valor_compra.Caption) + fValidaValor2(txt_valor_desconto.Text)
        MovDocEletronicoCabecalho.EtapaConcluida = Val(ETAPA_CONCLUIDA.PRE_PROCESSADO)
        
        'Calcula base de calculo do ICMS através do itens já adicionados
        AtualizaBaseCalculoICMSCabecalho
        'Call TestaConexao(11, "IniciaProcessoNFCe")
        
        'Alex - NFCE
'        If Val(txt_numero_nota_abastecimento.Text) > 0 Then
'            MovCupomFiscal.NumeroCheque = CLng(txt_numero_nota_abastecimento.Text)
'        Else
'            MovCupomFiscal.NumeroCheque = 0
'        End If


'        If i = 1 Then
'            MovCupomFiscal.ValorDesconto = fValidaValor2(txt_valor_desconto.Text)
'        End If
        
        
        If MovCupomFiscal.AlterarFormaPagamento(g_empresa, lNumeroPDV, lNumeroCupom, lData) Then
            'Call TestaConexao(12, "IniciaProcessoNFCe")
        'Teste quando o cupom tem iten(s) cancelado(s) e o total ficou zerado
            If fValidaValor(txt_valor_recebido.Text) = 0 And fValidaValor(lbl_valor_compra.Caption) = 0 Then
                If Not MovCupomFiscal.CancelaCupom(g_empresa, lNumeroPDV, lNumeroCupom, lData) Then
                    MsgBox "Não foi possível alterar cupom para cancelado!", vbInformation, "Erro de Integridade"
                End If
            End If
            If fValidaValor(txt_valor_desconto.Text) > 0 Then
                If MovCupomFiscal.AlterarDesconto(g_empresa, lNumeroPDV, lNumeroCupom, lData, lTotalCupom, fValidaValor(txt_valor_desconto.Text)) Then
                    If Not MovCupomFiscalItem.AlterarDesconto(g_empresa, lNumeroPDV, lNumeroCupom, lData, lTotalCupom, fValidaValor(txt_valor_desconto.Text)) Then
                        MsgBox "Não foi possível alterar o desconto no item de cupom!", vbInformation, "Erro de Integridade"
                    End If
                Else
                    MsgBox "Não foi possível alterar o desconto do cupom!", vbInformation, "Erro de Integridade"
                End If
            End If
            
            'se a forma de pagamento for dinheiro chama a função para incluir o registro no caixa de pista
            If lGeraCaixaDinheiro = True And cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 1 Then
                If Not IntegracaoCaixa.LocalizarNome(g_empresa, "DINHEIRO") Then
                    Call CriaLogCupom("Erro cmdIniciaProcessoNFCe:Integração de caixa inexistente. DINHEIRO")
                    MsgBox "Não será possível integrar com o caixa!", vbInformation, "Erro de Integridade"
                Else
                    If IncluiMovimentoCaixa(MovCupomFiscal.Data, MovCupomFiscal.Periodo, False, "DINHEIRO", fValidaValor(txt_valor_recebido.Text), "", "CF:" & MovCupomFiscal.NumeroCupom) Then

                    Else
                        Call CriaLogCupom("Erro cmdIniciaProcessoNFCe:Não integrada no caixa. DINHEIRO")
                        MsgBox "Não foi possível integrar no caixa!", vbInformation, "Erro de Integridade!"
                    End If
                End If
            End If
            
            'se a forma de pagamento for cheque a vista chama a função para incluir o registro no caixa de pista
            If lGeraCaixaChequeAVista = True And cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 2 Then
                If Not IntegracaoCaixa.LocalizarNome(g_empresa, "CHEQUE A VISTA") Then
                    Call CriaLogCupom("Erro cmdIniciaProcessoNFCe:Integração de caixa inexistente. CHEQUE A VISTA")
                    MsgBox "Não será possível integrar com o caixa!", vbInformation, "Erro de Integridade"
                Else
                    If IncluiMovimentoCaixa(MovCupomFiscal.Data, MovCupomFiscal.Periodo, False, "CHEQUE A VISTA", MovCupomFiscal.ValorTotal, "", "CF:" & MovCupomFiscal.NumeroCupom) Then

                    Else
                        Call CriaLogCupom("Erro cmdIniciaProcessoNFCe:Não integrada no caixa. CHEQUE A VISTA")
                        MsgBox "Não foi possível integrar no caixa!", vbInformation, "Erro de Integridade!"
                    End If
                End If
            End If

            If cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 5 Then
                IncluiNotaAbastecimento
            End If
            
            If cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 16 Then
            
                If lIntegraDescontoCartaoEspecial = True And lValorDescontoConcedido > 0 Then
                    If Not IncluiMovimentoCaixa(lDataCupom, lPeriodo, True, "DescontoCartaoEspecial", lValorDescontoConcedido, "", "Desconto Cartão Especial") Then
                        Call CriaLogCupom("Erro IniciaProcessamentoNFCe:Desconto Cartao Desconto Especial não integrada no caixa.")
                        MsgBox "Não foi possível integrar Desconto Cartão Especial no caixa!", vbInformation, "Erro de Integridade!"
                    End If
                End If

            End If

            If Not MovDocEletronicoCabecalho.AlterarInformacoesEtapaPagamento(g_empresa, lData, False, True, MODELO_NFCE, lSerieNFCe, lNumeroCupom) Then
                Call CriaLogCupom("Não foi possível alterar a forma de pagamento do documento eletrônico!")
                Call CriaLogCupom("Caso o possua mais de 1 item o valor total do cabecalho ficará incorreto")
                MsgBox "Não foi possível alterar a forma de pagamento do documento eletrônico!", vbInformation, "Erro de Integridade"
            End If
            If Not MovDocEletronicoItem.AlterarEtapaConcluida(g_empresa, lData, False, True, MODELO_NFCE, lSerieNFCe, lNumeroCupom, Val(ETAPA_CONCLUIDA.PRE_PROCESSADO)) Then
                Call CriaLogCupom("Não foi possível alterar a EtapaConcluída do Item do documento eletrônico!")
                MsgBox "Não foi possível alterar a EtapaConcluída do Item do documento eletrônico!", vbInformation, "Erro de Integridade"
            End If
            'Call TestaConexao(13, "IniciaProcessoNFCe")
            
            'Corrige Problema com o IdClienteFornecedor
            If Not MovDocEletronicoCabecalho.AlterarIdCliente(g_empresa, lData, False, True, MODELO_NFCE, lSerieNFCe, lNumeroCupom, l_codigo_cliente) Then
                Call CriaLogCupom("Não foi possível alterar o IdClienteFornecedor de documento eletrônico Cabecalho!")
                MsgBox "Não foi possível alterar o IdClienteFornecedor de documento eletrônico Cabecalho!", vbInformation, "Erro de Integridade"
            End If
            If Not MovDocEletronicoItem.AlterarIdCliente(g_empresa, lData, False, True, MODELO_NFCE, lSerieNFCe, lNumeroCupom, l_codigo_cliente) Then
                Call CriaLogCupom("Não foi possível alterar o IdClienteFornecedor de documento eletrônico Ítem!")
                MsgBox "Não foi possível alterar o IdClienteFornecedor de documento eletrônico Ítem!", vbInformation, "Erro de Integridade"
            End If
            
            If fValidaValor(txt_valor_desconto.Text) > 0 Then
                If Not MovDocEletronicoItem.AlterarDesconto(g_empresa, lData, False, True, MODELO_NFCE, lSerieNFCe, lNumeroCupom, lTotalCupom, fValidaValor(txt_valor_desconto.Text)) Then
                    MsgBox "Não foi possível alterar o desconto de ítens do documento eletrônico!", vbInformation, "Erro de Integridade"
                End If
            End If
            'Corrige o problemas dos valores da tributação no cabeçalho estarem indo zeradas
            If Not MovDocEletronicoCabecalho.AlterarValoresTributacaoCabecalho(g_empresa, lData, False, True, MODELO_NFCE, lSerieNFCe, lNumeroCupom) Then
                Call CriaLogCupom("Não foi possível alterar os totais das tributações do documento eletrônico Cabecalho!")
                Call CriaLogCupom("Dados: " & "Empresa: " & g_empresa & " Data: " & lData & " modelo:  " & MODELO_NFCE & "Série: " & lSerieNFCe & " Numero: " & lNumeroCupom)
                'MsgBox "Não foi possível alterar os totais das tributações do documento eletrônico Cabecalho!", vbInformation, "Erro de Integridade"
            End If
            
        Else
            Call CriaLogCupom("Não foi possível alterar a forma de pagamento")
            Call CriaLogCupom("Caso o possua mais de 1 item o valor total do cabecalho ficará incorreto")
            MsgBox "Não foi possível alterar a forma de pagamento!", vbInformation, "Erro de Integridade"
        End If
        If cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 5 Then
            VerificaImprimeNotaAbastecimento
        End If
        lDescontoEspecial = 0
        
        'PONTO DE INSERÇÃO #NFCE-001 - ALEX - NFCE
        'Call TestaConexao(14, "IniciaProcessoNFCe")
        Call EnviaDadosParaNFCe(lNumeroCupom, lData)
        
        l_flag_cupom_fiscal = "F"
        If lNotificacaoGic Then
            menu_personalizado.AtivaVerificacaoGIC
        End If
        cmdIniciaProcessoFinalizacaoNFCe.Enabled = False
        lCodigoFiscal = "  "
        btnDescontoCartaoEspecial.Enabled = False
        'mnuLeituraX.Enabled = True
        mnuPontoFuncionario.Enabled = True
        frm_fechamento_cupom.Width = 100
        frm_fechamento_cupom.Height = 100
        frm_fechamento_cupom.Top = 100
        frm_fechamento_cupom.Left = 100
        frm_fechamento_cupom.ZOrder 1
        frm_fechamento_cupom.Enabled = False
        Call MontaCupomVideo(lNumeroCupom, lData)
        Call BuscaRegistro(lNumeroCupom, lData, lOrdem - 1)
        'If MovCupomFiscal.FormaPagamento = 2 Or MovCupomFiscal.FormaPagamento = 3 Then
        '    MsgBox "Aguarde o final da impressão!" & Chr(10) & Chr(10) & "Coloque o cheque na impressora fiscal, tecle enter e aguarde.", vbExclamation, "Autenticação de Cheque"
        '    If lExisteImpressora Then
        '        xString = "001,002,004,008,016,032,064,128,064,016,008,004,002,001,129,129,129,129"
        '        BemaRetorno = Bematech_FI_ProgramaCaracterAutenticacao(xString)
        '        BemaRetorno = Bematech_FI_Autenticacao
        '    End If
        'End If
        'NovoCupom
        'mnuSenha_Click
        lOrdem = 0
        NovoCupom
        
        'mnuSenha_Click
        
        
        'alterar para posto ventania, para o usuario do caixa nao precisar ficar informando usuario e senha a todo cupom
        'aqui ao inves de chamar linha acima
        'faria algo para executar os comentarios abaixo
'                Call AbilitaMenu(True)
'                txt_cupom_fiscal.Enabled = True
'                txt_cliente = "0"
'                txt_cliente.SetFocus
        
        
        'cmd_bico(0).SetFocus
        AguardaProcessamentoNFCe (MovSolicitacaoFuncaoNFe.NSU_MovSolicitacaoFuncaoNFe)
        'Call TestaConexao(15, "IniciaProcessoNFCe")
        
        
        mnuSenha_Click

        
        'Teste pra tentar resolver problema que fecha conexao ao passar cartao pela seguna vez
        'Call CriaLogCupom("TestaConexao - Chama FinalizaConexao")
        'Conectar.FinalizaConexao
        'Call CriaLogCupom("TestaConexao - Concluida FinalizaConexao")
        '
        'Call CriaLogCupom("TestaConexao - Chama AbreConexao")
        'Conectar.AbreConexao
        'Call CriaLogCupom("TestaConexao - Concluida AbreConexao")
        'Unload Me
        'Screen.MousePointer = 1
    End If
End Sub
Private Function IntegraCartaoCreditoNoCaixa() As Boolean
    Dim xArqTxt As New FileSystemObject
    Dim xArquivo As TextStream
    Dim xNomeArquivo As String
    Dim xNomeArquivoCopia As String
    Dim xString As String
    Dim xNomeBandeira As String
    Dim xNomeBandeiraLido As String
    Dim xOperacao As String
    Dim xNumLinha As Integer
    Dim xPlanoValeCard As String
    Dim xQtdParcela As Integer
    Dim i As Integer
    Dim xMensagem As String
    Dim xDiferenciarCieloRedecard As Boolean
    Dim xNomeAdm As String
    
    On Error GoTo FileError
    
    xNomeBandeira = ""
    xNomeBandeiraLido = ""
    xOperacao = ""
    xPlanoValeCard = ""
    lCodigoCartao = 0
    xNumLinha = 0
    xQtdParcela = 0
    
    IntegraCartaoCreditoNoCaixa = False
    
    xDiferenciarCieloRedecard = True
    If ConfiguracaoDiversa.LocalizarCodigo(1, "CARTAO: DIFERENCIAR CIELO/REDECARD") Then
        xDiferenciarCieloRedecard = ConfiguracaoDiversa.Verdadeiro
    End If

    'xNomeArquivo = "c:\vb5\sgp\data\teste.txt"
    xNomeArquivo = "c:\vb5\sgp\data\teste.txt"
    xNomeArquivoCopia = "TTF_" & Format(Date, "dd") & "_" & Format(Date, "MM") & "_" & Format(Date, "yyyy") & "__" & Format(Now, "HH:mm:ss") & ".LOG"
    Mid(xNomeArquivoCopia, 19, 1) = "_"
    Mid(xNomeArquivoCopia, 22, 1) = "_"
    xNomeArquivoCopia = "c:\vb5\sgp\data\" & xNomeArquivoCopia
    If xArqTxt.FileExists(xNomeArquivo) Then
    
        Set xArquivo = xArqTxt.OpenTextFile(xNomeArquivo, ForReading)
        Do Until xArquivo.AtEndOfStream
            xString = xArquivo.ReadLine
            xNumLinha = xNumLinha + 1
            If Val(Mid(xString, 1, 3)) > 10 Or xDiferenciarCieloRedecard = False Then
                'Detecta venda parcelada
                If Mid(xString, 1, 7) = "018-000" Then
                    'xQtdParcelas = Mid(xString, 11, Len(xString) - 10)
                    xOperacao = "CREDITO"
                End If
                'Detecta a Administradora de Cartão
                If xNomeAdm = "" Then
                    If xString Like "*CIELO*" Then
                        xNomeAdm = "CIELO"
                    ElseIf xString Like "*REDECARD*" Or xString Like "*REDE*" Then
                        xNomeAdm = "REDECARD"
                    ElseIf xString Like "*SAFRAPAY*" Then
                        xNomeAdm = "SAFRAPAY"
                    End If
                ElseIf xNomeAdm = "REDECARD" Then
                    If UCase(xString) Like "*DEBITO*" Then
                        xOperacao = "DEBITO"
                        If xNomeBandeiraLido <> "" Then
                           xNomeBandeira = xNomeBandeiraLido & " " & xOperacao
                           Exit Do
                        End If
                        
                    ElseIf UCase(xString) Like "*CREDITO*" Then
                        xOperacao = "CREDITO"
                        If xNomeBandeiraLido <> "" Then
                           xNomeBandeira = xNomeBandeiraLido & " " & xOperacao
                           Exit Do
                        End If
                    End If
                End If
                'Detecta a Bandeira
                If xNomeBandeira = "" Then
' ELECTRON CRIADO por tasso em 03/03/2017 pra resolver problema no VILA ALZIRA
                    If xString Like "*ELECTRON*" Then
                        xNomeBandeira = "VISA DEBITO"
                        xOperacao = "DEBITO"
                        Exit Do
                    ElseIf xString Like "*MASTER*" And xNomeAdm = "SAFRAPAY" Then
                        xNomeBandeiraLido = "MASTERCARD"
                        xNomeBandeira = "MASTERCARD"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*MAESTRO*" And xNomeAdm = "SAFRAPAY" Then
                        xNomeBandeiraLido = "MAESTRO"
                        xNomeBandeira = "MAESTRO"
                        xOperacao = "DEBITO"
                        Exit Do
                    ElseIf xString Like "*VISADEB*" And xNomeAdm = "SAFRAPAY" Then
                        xNomeBandeira = "VISA DEBITO"
                        xOperacao = "DEBITO"
                        Exit Do
                    ElseIf xString Like "*VISACRE*" And xNomeAdm = "SAFRAPAY" Then
                        xNomeBandeira = "VISA CREDITO"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*ELO DEB*" And xNomeAdm = "SAFRAPAY" Then
                        xNomeBandeiraLido = "ELO"
                        xNomeBandeira = "ELO DEBITO"
                        xOperacao = "DEBITO"
                        Exit Do
                    ElseIf xString Like "*ELO CRE*" And xNomeAdm = "SAFRAPAY" Then
                        xNomeBandeiraLido = "ELO"
                        xNomeBandeira = "ELO CREDITO"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*ELOCRE*" And xNomeAdm = "SAFRAPAY" Then
                        xNomeBandeiraLido = "ELO"
                        xNomeBandeira = "ELO DEBITO"
                        xOperacao = "DEBITO"
                        Exit Do
                    ElseIf xNomeAdm = "REDECARD" And xString Like "*ALELO*" Then
                        xNomeBandeiraLido = "ALELO"
                    ElseIf xNomeAdm = "REDECARD" And xString Like "*ELO*" Then
                        xNomeBandeiraLido = "ELO"
                        
' MAESTRO CRIADO por tasso em 03/03/2017 pra resolver problema no VILA ALZIRA
                    ElseIf xString Like "*MAESTRO*" Then
                        xNomeBandeira = "MAESTRO"
                        xOperacao = "DEBITO"
                        Exit Do
                    ElseIf xString Like "*FLEX CAR VISA VALE*" Then
                        xNomeBandeira = "FLEX CAR VISA"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*FLEX CAR VISA VALE*" Then
                        xNomeBandeira = "FLEX CAR VISA"
                        xOperacao = "DEBITO"
                        Exit Do
                    ElseIf xString Like "*VISA*" Then
                        xNomeBandeiraLido = "VISA"
                        xNomeBandeira = "VISA"
                    ElseIf xString Like "*VISANET*" Then
                        xNomeBandeiraLido = "VISANET"
                        xNomeBandeira = "VISA"
'                    ElseIf xString Like "*REDECARD*" Then
'                        xNomeBandeiraLido = "REDECARD"
'                        xNomeBandeira = "REDECARD"
                    ElseIf xString Like "*MASTERCARD*" Then
                        If xNomeAdm <> "" Then
                            xNomeBandeiraLido = "REDECARD"
                            xNomeBandeira = "REDECARD"
                            If xOperacao = "" Then
                                xOperacao = "CREDITO"
                            End If
                            Exit Do
                        Else
                            xNomeBandeiraLido = "MASTERCARD"
                            xNomeBandeira = "REDECARD"
                        End If
' comentado por tasso em 03/03/2017 pra resolver problema no VILA ALZIRA
'                    ElseIf xString Like "*MAESTRO*" Then
'                        'xNomeBandeiraLido = "MASTERCARD"  ' 07/05/2015
'                        xNomeBandeiraLido = "MAESTRO"      ' 07/05/2015
'                        xNomeBandeira = "REDECARD"
                    ElseIf xString Like "*AMEX*" Then
                        xNomeBandeiraLido = "AMERICAN EXPRESS"
                        xNomeBandeira = "AMERICAN EXPRESS"
                        xOperacao = "CRED"
                        Exit Do
                    ElseIf xString Like "*AMERICAN EXPRESS*" Then
                        xNomeBandeiraLido = "AMERICAN EXPRESS"
                        xNomeBandeira = "AMERICAN EXPRESS"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*ELO CREDITO*" Then
                        xNomeBandeiraLido = "ELO"
                        xNomeBandeira = "ELO CREDITO"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*ELO DEBITO*" Or xString Like "*ELO_DEB*" Then
                        xNomeBandeiraLido = "ELO"
                        xNomeBandeira = "ELO DEBITO"
                        xOperacao = "DEBITO"
                        Exit Do
                    ElseIf xString Like "*SODEXO MOBILITY*" Then
                        xNomeBandeiraLido = "SODEXO"
                        xNomeBandeira = "SODEXO DEBITO"
                        xOperacao = "DEBITO"
                        Exit Do
                    ElseIf xString Like "*SODEXO*" Then
                        xNomeBandeiraLido = "SODEXO"
                        xNomeBandeira = "SODEXO CREDITO"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*BRASIL CARD*" Then
                        xNomeBandeira = "BRASIL CARD CREDITO"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*USACARDFROTA*" Then
                        xNomeBandeira = "USA CARD FROTAS CREDITO"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*USA CARD*" Or xString Like "*USACARD*" Then
                        xNomeBandeira = "USA CARD CREDITO"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*LOSANGO*" Then
                        xNomeBandeira = "PETROBRAS CREDITO"
                        Exit Do
                    ElseIf xString Like "*CHEQUE ELETRONICO*" Then
                        xNomeBandeira = "PRE DATADO 45 DIAS CREDITO"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*TICKETLOG*" Then
                        xNomeBandeiraLido = "TICKETLOG"
                        xNomeBandeira = "TICKETLOG"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*ACCOR*" Or xString Like "*TICKET*" Then
                        xNomeBandeira = "TICKET CAR SMART CREDITO"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*SERVICO: SERV*" Then
                        xNomeBandeira = "TICKET CAR SMART CREDITO"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*G O O D C A R D*" Then
                        xNomeBandeira = "GOOD CARD CREDITO"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*GETNET*" Then
                        If g_nome_empresa Like "*MARTINS E BISPO*" Or g_nome_empresa Like "*POSTO SOLEX*" Then
                            xNomeAdm = "GETNET"
                            xNomeBandeira = "GETNET"
                        Else
                            xNomeBandeira = "GOOD CARD CREDITO"
                            xOperacao = "CREDITO"
                        End If
                        'Exit Do
                    ElseIf xString Like "*HIPERCARD*" Then
                        xNomeBandeira = "HIPERCARD"
                        xOperacao = "CREDITO"
                        Exit Do
                    'NEW 03/11/15 DAQUI ATE..
                    'cartao hipecard nao esta caindo no sistema (posto rubi)
                    'obs: ele passa somente na adm redecard
                    ElseIf xString Like "*HIPERCARD*" Or xString Like "*HIPER*" Then
                        xNomeBandeiraLido = "HIPERCARD"
                        xNomeBandeira = "HIPERCARD CREDITO"
                        xOperacao = "CREDITO"
                        Exit Do
                    '... aqui new 03/11/2015
                    ElseIf xString Like "*VALECARD*" Then
                        xNomeBandeira = "VALECARD"
                        xOperacao = "CREDITO"
                    ElseIf xString Like "*SOROCRED*" Then 'SOROCRED
                        xNomeBandeira = "SOROCRED"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*AGIPLAN*" Then 'AGIPLAN
                        xNomeBandeira = "AGIPLAN"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*PAGCARD*" Then
                        xNomeBandeira = "PAGCARD CREDITO"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*R E D E    R A T I N H O*" Then
                        xNomeBandeira = "REDE RATINHO CREDITO"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*R E D E    V W*" Then
                        xNomeBandeira = "REDE VW CREDITO"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*VALE SHOP*" Then
                        xNomeBandeira = "VALE SHOP"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*PRIVATE LABEL*" Then
                        xNomeBandeiraLido = "PRIVATE LABEL"
                        xNomeBandeira = "PRIVATE LABEL CREDITO"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*VISA CREDITO*" Then
                        xNomeBandeira = "VISA CREDITO"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*POLICARD*" Or xString Like "*PREMIACAO*" Or xString Like "*CONVENIO*" Then
                        xNomeBandeiraLido = "POLICARD"
                        xNomeBandeira = "POLICARD"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*Policard*" Then
                        xNomeBandeiraLido = "POLICARD"
                        xNomeBandeira = "POLICARD"
                        xOperacao = "CREDITO"
                        Exit Do
                    'ElseIf g_nome_empresa Like "*VALPOSTO*" And xString Like "*PRIV LABEL*" And xNomeBandeiraLido = "REDECARD" And xNomeBandeira = "REDECARD" Then
                    ElseIf g_nome_empresa Like "*VALPOSTO*" And xString Like "*PRIV LABEL*" And xNomeBandeira = "REDECARD" Then
                        xNomeBandeiraLido = "IPIRANGA CREDITO"
                        xNomeBandeira = "IPIRANGA CREDITO"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*CABAL DEBITO*" Then
                        xNomeBandeira = "CABAL DEBITO"
                        xOperacao = "DEBITO"
                        Exit Do
                    ElseIf xString Like "*CABAL CREDITO*" Then
                        xNomeBandeira = "CABAL CREDITO"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*CABAL*" Then
                        xNomeBandeiraLido = "CABAL"
                        xNomeBandeira = "CABAL"
                    ElseIf xString Like "*CREDSYSTEM*" Then
                        xNomeBandeira = "CREDSYSTEM CREDITO"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*DINERS*" Then
                        xNomeBandeira = "DINERS CREDITO"
                        xOperacao = "CREDITO"
                        Exit Do
                        'NEW 06/07/2015 DAQUI ATE...
                    ElseIf xString Like "*DISCOVER*" Then
                        xNomeBandeira = "CIELO DISCOVER"
                        xOperacao = "CREDITO"
                        Exit Do
                        'NEW 19/08/2015
                    ElseIf xString Like "*GOODCARD*" Then
                        xNomeBandeira = "GOODCARD CREDITO"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*VERYCARD*" Then
                        xNomeBandeira = "SENACARD"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*UP CREDITO*" Then
                        xNomeBandeira = "UP"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*UP DEBITO*" Then
                        xNomeBandeira = "UP"
                        xOperacao = "DEBITO"
                        Exit Do
                    ElseIf xString Like "*UP BRASIL*" Then
                        xNomeBandeira = "UP"
                        xOperacao = "CREDITO"
                        Exit Do
                        'NEW 19/08/2015
                        'NEW 06/07/2015 ...AQUI
'                    ElseIf xString Like "*FLEX CAR VISA VALE*" Then
'                        xNomeBandeira = "FLEX CAR VISA"
'                        xOperacao = "CREDITO"
'                        Exit Do
'                    ElseIf xString Like "*FLEX CAR VISA VALE*" Then
'                        xNomeBandeira = "FLEX CAR VISA"
'                        xOperacao = "DEBITO"
'                        Exit Do
                    ElseIf xString Like "*FITCARD*" Then
                        xNomeBandeira = "FITCARD"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*ALELO REFEICAO*" Then
                        xNomeBandeira = "ALELO REFEICAO"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*ALELO BENEFICIOS*" Then
                        xNomeBandeira = "ALELO BENEFICIOS"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*VR BENEFICIOS*" Then
                        xNomeBandeira = "VR BENEFICIOS"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*CREDZ*" Then
                        xNomeBandeira = "CREDZ"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*MAIS*" Then
                        xNomeBandeira = "MAIS"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*ALELO*" Then
                        xNomeBandeiraLido = "ALELO"
                        xNomeBandeira = "ALELO"
                    ElseIf xString Like "*VR*" Then
                        xNomeBandeiraLido = "VR"
                        xNomeBandeira = "VR AUTO"
                        xOperacao = "CREDITO"
                        Exit Do
                    End If
                End If
                
                'IMPLEMENTAÇÃO EXCLUSIVA PARA "MARTINS E BISPO" E "*POSTO SOLEX*"
                If xNomeAdm = "GETNET" And xNomeBandeira = "GETNET" Then
                    If Mid(xString, 1, 7) = "029-009" Then
                        xNomeAdm = ""
                        If xString Like "*VISA CREDITO*" Then
                            xNomeBandeiraLido = "VISA CREDITO"
                            xNomeBandeira = "VISA"
                            xOperacao = "CREDITO"
                            Exit Do
                        ElseIf xString Like "*VISA DEBITO*" Then
                            xNomeBandeiraLido = "VISA DEBITO"
                            xNomeBandeira = "VISA"
                            xOperacao = "DEBITO"
                            Exit Do
                        End If
                    End If
                End If
                
                
                'Este teste esta fora dos testes anterior pelo motivo que a bandeira
                'estaria REDECARD e nao em branco.
                If xString Like "*IPIRANGA*" Or xString Like "*FININVEST*" Then 'FININVEST
                    xNomeBandeira = "IPIRANGA CREDITO"
                    xOperacao = "CREDITO"
                    Exit Do
                End If
                
                'Detecta se é Débito ou Crédito
                If xOperacao = "" Then
                    If xString Like "*CREDITO*" Or xString Like "*CRÉDITO*" Then
                        xOperacao = "CREDITO"
                        
                        If xNomeBandeiraLido <> "" And xNomeAdm = "REDECARD" Then
                            xNomeBandeira = xNomeBandeiraLido & " " & xOperacao
                        End If
                        
                        If xNomeBandeira <> "" Then
                            Exit Do
                        End If
                    End If
                    If xString Like "*DEBITO*" Or xString Like "*DÉBITO*" Or xString Like "*ELECTRON*" Then
                        xOperacao = "DEBITO"
                        
                        If xNomeBandeiraLido <> "" And xNomeAdm = "REDECARD" Then
                            xNomeBandeira = xNomeBandeiraLido & " " & xOperacao
                        End If
                        
                        If xNomeBandeira <> "" Then
                            Exit Do
                        End If
                    End If
                    
                    If xNomeBandeira = "REDECARD" And xNomeAdm = "CIELO" And xString Like "*MAESTRO*" Then
                        xOperacao = "DEBITO"
                        Exit Do
                    ElseIf xNomeBandeira = "REDECARD" And xNomeAdm = "CIELO" And xString Like "*MASTERCARD*" Then
                        xOperacao = "CREDITO"
                        Exit Do
                    End If

                    
                End If
                
                'Detecta Plano do ValeCard
                If xNomeBandeira = "VALECARD" And xPlanoValeCard = "" Then
                    If xString Like "*PLANO-*" Then
                        For i = 1 To (Len(xString) - 8)
                            If Mid(xString, i, 6) = "PLANO-" Then
                                xOperacao = "CREDITO"
                                xPlanoValeCard = "PLANO-" & Mid(xString, i + 6, 1)
                                Exit Do
                                Exit For
                            End If
                        Next
                    End If
                End If
                
                'Detecta qtd Parcelas REDECARD
                If xNomeBandeira = "REDECARD" And xOperacao = "" Then
                    If xString Like "*NUMERO DE PARCELAS*" Then
                        xString = Mid(xString, 12, Len(xString) - 12)
                        For i = 1 To Len(xString)
                            If IsNumeric(Mid(xString, i, 1)) Then
                                xOperacao = xOperacao & Mid(xString, i, 1)
                            End If
                        Next
                        If IsNumeric(xOperacao) Then
                            xQtdParcela = Val(xOperacao)
                            xOperacao = "CREDITO"
                            Exit Do
                        End If
                    End If
                End If
                
                
            End If
        Loop
        xArquivo.Close
        
        CriaLogECF (Date & " " & Time & "[IntegraCartaoCreditoNoCaixa] - Leitura do arquivo finalizada: xNomeBandeira = " & xNomeBandeira & " - xOperacao=" & xOperacao & " - xNomeAdm=" & xNomeAdm)
       
        'Quando for CIELO
        'Chega aqui com xNomeBandeira = "REDECARD"
        'xOperacao = ""
        'Se for xNomeBandeiraLido = "MASTERCARD"  entao será "CREDITO"
        'Se for xNomeBandeiraLido = "MAESTRO"     entao será "DEDITO"
        If xNomeBandeira = "REDECARD" And xOperacao = "" Then
            If xNomeBandeiraLido = "MASTERCARD" Then
                xOperacao = "CREDITO"
            ElseIf xNomeBandeiraLido = "MAESTRO" Then
                xOperacao = "DEBITO"
            Else
                xOperacao = "CREDITO"
            End If
        End If
       
       'Teste cartao: ao chegar neste ponto mudar o valor da variavel g_nome_empresa para "*POSTO T13*"
        If g_nome_empresa Like "*POSTO T13*" Or g_nome_empresa Like "*MARQUES DE CASTRO*" Or g_nome_empresa Like "*AUTO POSTO CLASSE A*" Or g_nome_empresa Like "*FLORIDIAN COMERCIO E PARTICIPAÇOES*" Or g_nome_empresa Like "*YPE COM DE COMB E LUBRIFICANTES*" Then
            If xNomeBandeira = "REDECARD" And xOperacao = "DEBITO" Then
                xNomeBandeira = "MAESTRO"
            ElseIf xNomeBandeira = "REDECARD" And xOperacao = "CREDITO" Then
                xNomeBandeira = "MASTERCARD"
            End If
        End If
       
        If xNomeBandeira <> "" And xOperacao <> "" Then
            If CartaoCredito.LocalizarPrimeiro Then
                If xNomeBandeira = "VALECARD" Then
                    If UCase(CartaoCredito.Nome) Like "*" & xNomeBandeira & "*" Then
                        If UCase(CartaoCredito.Nome) Like "*" & xOperacao & "*" Then
                            lCodigoCartao = CartaoCredito.Codigo
                            If UCase(CartaoCredito.Nome) Like "*" & xPlanoValeCard & "*" Then
                                lCodigoCartao = CartaoCredito.Codigo
                            End If
                        End If
                    End If
                Else
                    If g_nome_empresa Like "*POSTO T13*" Or g_nome_empresa Like "*MARQUES DE CASTRO*" Or g_nome_empresa Like "*AUTO POSTO CLASSE A*" Or g_nome_empresa Like "*FLORIDIAN COMERCIO E PARTICIPAÇOES*" Or g_nome_empresa Like "*YPE COM DE COMB E LUBRIFICANTES*" Then
                        If UCase(CartaoCredito.Nome) Like "*" & xNomeBandeira & "*" Then
                            If UCase(CartaoCredito.Nome) Like "*" & xOperacao & "*" Then
                                'If xNomeBandeira Like "*MAESTRO*" Or xNomeBandeira Like "*MASTERCARD*" Or xNomeBandeira Like "*VISA*" Then
                                    If UCase(CartaoCredito.Nome) Like "*" & xNomeAdm & "*" Then
                                        lCodigoCartao = CartaoCredito.Codigo
                                    End If
                                'Else
                                    'lCodigoCartao = CartaoCredito.Codigo
                                'End If
                            End If
                        End If
                    Else
                        If UCase(CartaoCredito.Nome) Like "*" & xNomeBandeira & "*" Then
                            If UCase(CartaoCredito.Nome) Like "*" & xOperacao & "*" Then
                                lCodigoCartao = CartaoCredito.Codigo
                            End If
                        End If
                    End If
                End If
                If lCodigoCartao = 0 Then
                    Do Until CartaoCredito.LocalizarProximo = False
                        If g_nome_empresa Like "*POSTO T13*" Or g_nome_empresa Like "*MARQUES DE CASTRO*" Or g_nome_empresa Like "*AUTO POSTO CLASSE A*" Or g_nome_empresa Like "*FLORIDIAN COMERCIO E PARTICIPAÇOES*" Or g_nome_empresa Like "*YPE COM DE COMB E LUBRIFICANTES*" Then
                            If UCase(CartaoCredito.Nome) Like "*" & xNomeBandeira & "*" Then
                                If UCase(CartaoCredito.Nome) Like "*" & xOperacao & "*" Then
                                    'If xNomeBandeira = "MAESTRO" Or xNomeBandeira = "MASTERCARD" Or xNomeBandeira = "VISA" Then
                                        If UCase(CartaoCredito.Nome) Like "*" & xNomeAdm & "*" Then
                                            lCodigoCartao = CartaoCredito.Codigo
                                            Exit Do
                                        End If
                                    'Else
                                    '    lCodigoCartao = CartaoCredito.Codigo
                                    '    Exit Do
                                    'End If
                                End If
                            End If
                        Else
                            If UCase(CartaoCredito.Nome) Like "*" & xNomeBandeira & "*" Then
                                If UCase(CartaoCredito.Nome) Like "*" & xOperacao & "*" Then
                                    lCodigoCartao = CartaoCredito.Codigo
                                    Exit Do
                                End If
                                
                            ElseIf UCase(CartaoCredito.Nome) Like "*REDECARD*" And (xNomeBandeira Like "*MAESTRO*" Or xNomeBandeira Like "*MASTERCARD*") Then
                                If UCase(CartaoCredito.Nome) Like "*" & xOperacao & "*" Then
                                    lCodigoCartao = CartaoCredito.Codigo
                                    Exit Do
                                End If
                            End If
                        End If
                    Loop
                End If
            End If
        End If
        If lCodigoCartao > 0 Then
            Call CriaLogECF(Date & " " & Time & " IntegraCartaoCreditoNoCaixa: 1 lCodigoCartao=" & lCodigoCartao & " - xNomeBandeira=" & xNomeBandeira & " - xNomeAdm=" & xNomeAdm & " - xNomeArquivo=" & xNomeArquivo)
            Call BuscaNsuCartaoCredito(xNomeAdm, xNomeBandeira, xNomeArquivo)
            If lCartaoAutorizacao = "" And lCartaoNSU = "" Then
                Call BuscaDadosCartaoTefCerrado(xNomeArquivo)
                
                Call CriaLogECF(Date & " " & Time & " IntegraCartaoCreditoNoCaixa: 1 lCodigoCartao=" & lCodigoCartao & "lCartaoAutorizacao=" & lCartaoAutorizacao & " - lCartaoNSU=" & lCartaoNSU)
            End If
            Call CriaLogECF(Date & " " & Time & " IntegraCartaoCreditoNoCaixa: 2")
            IntegraCartaoCreditoNoCaixa = True
            'cópia para Testar Autorizacao e NSU
            'Call xArqTxt.CopyFile(xNomeArquivo, xNomeArquivoCopia, True)
            'Deleta Arquivo "c:\vb5\sgp\data\teste.txt"
            Call xArqTxt.DeleteFile(xNomeArquivo, True)
        Else
            Call GravaAuditoria(1, Me.name, 26, "Cartão Não Integrado: " & xNomeBandeira & " Operação: " & xOperacao & " Valor: " & lValorTotalUltimoCupom & " Linha: " & xNumLinha)
'            gNumeroEmailInicial = 0
'
'            xMensagem = "Empresa: " & g_nome_empresa & vbCrLf
'            xMensagem = xMensagem & "Data: " & Format(Date, "dd/mm/yyyy") & " as " & Format(Time, "HH:MM:SS") & vbCrLf & vbCrLf
'            xMensagem = xMensagem & "Cartao Nao Integrado em:" & Format(Date, "dd/mm/yyyy") & " " & Format(Time, "HH:MM:SS") & vbCrLf
'            xMensagem = xMensagem & " Bandeira: " & xNomeBandeira & " Operacao: " & xOperacao & " Valor: " & lValorTotalUltimoCupom & " Linha: " & xNumLinha & vbCrLf
'            xMensagem = xMensagem & "xNomeArquivo:" & xNomeArquivo & vbCrLf
'            xMensagem = xMensagem & "xNomeArquivoCopia:" & xNomeArquivoCopia & vbCrLf
'            Call EnviaMensagemEmail(g_empresa, g_nome_empresa, "Cartao Nao Integrado!", xMensagem, True, gNumeroEmailInicial)
            'tira cópia do arquivo "c:\vb5\sgp\data\teste.txt"
            'para o arquivo        "c:\vb5\sgp\data\TTF_dd_MM_yyyy__HH:mm:ss.LOG"
            Call xArqTxt.CopyFile(xNomeArquivo, xNomeArquivoCopia, True)
        End If
        Set xArquivo = Nothing
        Set xArqTxt = Nothing
    End If
    Exit Function

FileError:
    Call GravaAuditoria(1, Me.name, 26, "Cartão Não Integrado: " & Error & " Valor: " & lValorTotalUltimoCupom)
End Function
Private Sub BuscaNsuCartaoCredito(ByVal pNomeAdm As String, ByVal pNomeBandeira As String, ByVal pNomeArquivo As String)
    Dim xArqTxt As New FileSystemObject
    Dim xArquivo As TextStream
    Dim xString As String
    Dim xStringAutorizacao As String
    Dim xStringAutorizacao2 As String
    Dim xStringNSU As String
    Dim xTamanhoAutorizacao2 As Integer
    Dim xTamanhoAutorizacao As Integer
    Dim xTamanhoNSU As Integer
    Dim xIniciou As Boolean
    Dim i As Integer
    Dim i2 As Integer

    On Error GoTo FileError

    lCartaoAutorizacao = ""
    lCartaoNSU = ""
    lCartaoDataVencimento = "00:00:00"
    xStringAutorizacao = "TEXTO INEXISTENTE"
    xStringAutorizacao2 = "TEXTO INEXISTENTE"
    xStringNSU = "TEXTO INEXISTENTE"
    If xArqTxt.FileExists(pNomeArquivo) Then
        If pNomeBandeira = "AMERICAN EXPRESS" Then
            xStringAutorizacao = "AUTORIZ.="
            xStringNSU = "DOC="
        ElseIf pNomeBandeira = "GOOD CARD CREDITO" Then
            xStringAutorizacao = "Autorizacao :"
            xStringNSU = "KM          :"
        ElseIf pNomeBandeira = "REDECARD" Then
            xStringAutorizacao = "AUTO: "
            xStringAutorizacao2 = "AUTE: "
            xStringNSU = "CV: "
        ElseIf pNomeBandeira = "VISA" Then
            xStringAutorizacao = "AUT.="
            xStringNSU = "DOC=  "
        End If
        If pNomeAdm = "CIELO" Then
            xStringAutorizacao = "AUT="
            xStringAutorizacao2 = ""
            xStringNSU = "DOC="
        ElseIf pNomeAdm = "REDECARD" Then
            xStringAutorizacao = "AUTO:"
            xStringAutorizacao2 = ""
            xStringNSU = "CV:"
        End If
        xTamanhoAutorizacao = Len(xStringAutorizacao)
        xTamanhoAutorizacao2 = Len(xStringAutorizacao2)
        xTamanhoNSU = Len(xStringNSU)
        
        Set xArquivo = xArqTxt.OpenTextFile(pNomeArquivo, ForReading)
        Do Until xArquivo.AtEndOfStream
            xString = xArquivo.ReadLine
            If Mid(xString, 1, 7) = "012-000" Then
                lCartaoNSU = Mid(xString, 11, Len(xString) - 10)
            ElseIf Mid(xString, 1, 7) = "013-000" Then
                lCartaoAutorizacao = Mid(xString, 11, Len(xString) - 10)
                Exit Do
            End If
'            For i = 1 To Len(xString) - 1
'                'Testa se é Autorização
'                If Mid(xString, i, xTamanhoAutorizacao) = xStringAutorizacao Then
'                    xIniciou = False
'                    For i2 = i + xTamanhoAutorizacao To Len(xString) - 1
'                        If Mid(xString, i2, 1) <> " " Then
'                            If xIniciou = False Then
'                                xIniciou = True
'                                lCartaoAutorizacao = Mid(xString, i2, 1)
'                            Else
'                                lCartaoAutorizacao = lCartaoAutorizacao & Mid(xString, i2, 1)
'                            End If
'                        Else
'                            If xIniciou = True Then
'                                Exit For
'                            End If
'                        End If
'                    Next
'                'Testa se é Autorização
'                ElseIf Mid(xString, i, xTamanhoAutorizacao2) = xStringAutorizacao2 Then
'                    xIniciou = False
'                    For i2 = i + xTamanhoAutorizacao2 To Len(xString) - 1
'                        If Mid(xString, i2, 1) <> " " Then
'                            If xIniciou = False Then
'                                xIniciou = True
'                                lCartaoAutorizacao = Mid(xString, i2, 1)
'                            Else
'                                lCartaoAutorizacao = lCartaoAutorizacao & Mid(xString, i2, 1)
'                            End If
'                        Else
'                            If xIniciou = True Then
'                                Exit For
'                            End If
'                        End If
'                    Next
'                'Testa se é NSU
'                ElseIf Mid(xString, i, xTamanhoNSU) = xStringNSU Then
'                    xIniciou = False
'                    For i2 = i + xTamanhoNSU To Len(xString) - 1
'                        If Mid(xString, i2, 1) <> " " Then
'                            If xIniciou = False Then
'                                xIniciou = True
'                                lCartaoNSU = Mid(xString, i2, 1)
'                            Else
'                                lCartaoNSU = lCartaoNSU & Mid(xString, i2, 1)
'                            End If
'                        Else
'                            If xIniciou = True Then
'                                Exit For
'                            End If
'                        End If
'                    Next
'                End If
'            Next
'            If lCartaoAutorizacao <> "" And lCartaoNSU <> "" Then
'                Exit Do
'            End If
        Loop
        xArquivo.Close
    
        ' VERIFICA SE FOR FITCARD
        ' PARA VER SE É CARTAO DOS "CORREIOS GOIAS"
        If pNomeBandeira = "FITCARD" Then
            lIntegraDescontoCartaoCorreios = False
            Set xArquivo = xArqTxt.OpenTextFile(pNomeArquivo, ForReading)
            Do Until xArquivo.AtEndOfStream
                xString = xArquivo.ReadLine
                If xString Like "*CORREIOS GOIAS*" Then
                    lIntegraDescontoCartaoCorreios = True
                    Exit Do
                End If
            Loop
            xArquivo.Close
        End If
    
    End If
    Exit Sub

FileError:
    Call CriaLogECF(Date & " " & Time & " BuscaNsuCartaoCredito: " & Error)
End Sub
Private Sub BuscaDadosCartaoTefCerrado(ByVal pNomeArquivo As String)
    Dim xCampo010GP As String
    Dim xArqTxt As New FileSystemObject
    Dim xArquivo As TextStream
    Dim xString As String

    On Error GoTo FileError

    lCartaoAutorizacao = ""
    lCartaoNSU = ""
    lCartaoDataVencimento = "00:00:00"
    xCampo010GP = ""
    'Busca Data do Vencimento
    'Somente para TEFCERRADO
    If xArqTxt.FileExists(pNomeArquivo) Then
        Set xArquivo = xArqTxt.OpenTextFile(pNomeArquivo, ForReading)
        Do Until xArquivo.AtEndOfStream
            xString = xArquivo.ReadLine
            If xCampo010GP = "TEFCERRADO" Then
                If Mid(xString, 1, 7) = "012-000" Then
                    lCartaoNSU = Mid(xString, 11, Len(xString) - 10)
                ElseIf Mid(xString, 1, 7) = "013-000" Then
                    lCartaoAutorizacao = Mid(xString, 11, Len(xString) - 10)
                ElseIf Mid(xString, 1, 7) = "019-000" Then
                    lCartaoDataVencimento = Mid(xString, 11, Len(xString) - 10)
                    Exit Do
                End If
            Else
                If Mid(xString, 1, 7) = "010-000" Then
                    xCampo010GP = Mid(xString, 11, Len(xString) - 10)
                End If
            End If
        Loop
        xArquivo.Close
    End If
    Exit Sub
FileError:
    Call CriaLogECF(Date & " " & Time & " BuscaDadosCartaoTefCerrado: " & Error)
End Sub
Private Function LocalizarNCM(ByVal pTabela As Integer, ByVal pCodigo As String) As Boolean
    LocalizarNCM = False
    If Trim(pCodigo) = "" Then
        MsgBox "O produto " & Trim(Produto.Nome) & vbCrLf & "Está cadastrado sem NCM.", vbInformation, "Produto Sem NCM!"
        Exit Function
    End If
    If Not PercentualImposto.LocalizarCodigo(pTabela, pCodigo) Then
        MsgBox "O produto " & Trim(Produto.Nome) & vbCrLf & "De NCM:" & Trim(pCodigo) & vbCrLf & "Está sem NCM cadastrado.", vbInformation, "NCM Inexistente!"
        Exit Function
    End If
    LocalizarNCM = True
End Function
Private Sub LoopAbreAutomacao()
    Dim xSaiLoop As Boolean
    Dim xRetVal As Long
    
    xSaiLoop = False
    Do Until xSaiLoop = True
        If lMarcaAutomacao = "COMPANY" Then
            xRetVal = Shell("C:\Cerrado\AutoCerradoCompany\AutoCerradoCompany.EXE", vbMinimizedNoFocus)
            Call AguardaMS(2000)
            If ComunicaAutomacaoCerradoBD("AUTOMACAO ATIVADA", "") = True Then
                Call GravaAuditoria(1, Me.name, 26, "AutoCerradoCompany Aberto e Respondendo normalmente")
                xSaiLoop = True
            Else
                If (MsgBox("Deseja tentar abrir programa de comunicação com o equipamento de Automação novamente?", vbQuestion + vbYesNo + vbDefaultButton1, "Erro ao Abrir Progrma da Automação!") = vbNo) Then
                    Call GravaAuditoria(1, Me.name, 26, "Usuário desistiu de abrir AutoCerradoCompany")
                    xSaiLoop = True
                End If
            End If
        ElseIf lMarcaAutomacao = "HOROUSTECH" Then
            xRetVal = Shell("C:\Cerrado.Net\AutoCerradoHorousTech\AutoCerradoHorousTech.EXE", vbMinimizedNoFocus)
            Call AguardaMS(2000)
            If ComunicaAutomacaoCerradoBD("AUTOMACAO ATIVADA", "") = True Then
                Call GravaAuditoria(1, Me.name, 26, "AutoCerradoHorousTech Aberto e Respondendo normalmente")
                xSaiLoop = True
            Else
                If (MsgBox("Deseja tentar abrir programa de comunicação com o equipamento de Automação novamente?", vbQuestion + vbYesNo + vbDefaultButton1, "Erro ao Abrir Progrma da Automação!") = vbNo) Then
                    Call GravaAuditoria(1, Me.name, 26, "Usuário desistiu de abrir AutoCerradoHorousTech")
                    xSaiLoop = True
                End If
            End If
        ElseIf lMarcaAutomacao = "EZTECH" Then
            xRetVal = Shell("C:\Cerrado\AutoCerradoEZ\AutoCerradoEZ.EXE", vbMinimizedNoFocus)
            Call AguardaMS(2000)
            If ComunicaAutomacaoCerradoBD("AUTOMACAO ATIVADA", "") = True Then
                Call GravaAuditoria(1, Me.name, 26, "AutoCerradoEZ Aberto e Respondendo normalmente")
                xSaiLoop = True
            Else
                If (MsgBox("Deseja tentar abrir programa de comunicação com o equipamento de Automação novamente?", vbQuestion + vbYesNo + vbDefaultButton1, "Erro ao Abrir Progrma da Automação!") = vbNo) Then
                    Call GravaAuditoria(1, Me.name, 26, "Usuário desistiu de abrir AutoCerradoEZ")
                    xSaiLoop = True
                End If
            End If
        End If
    Loop
End Sub
Private Function LoopAbrePetromovelAuto(ByVal pExecutarComunicacaoBanco As Boolean) As Boolean
    Dim xSaiLoop As Boolean
    Dim xRetVal As Long
    Dim xCaminho As String
    Dim xCaminhoAmbienteDev As String
    

    LoopAbrePetromovelAuto = False
 
    xCaminhoAmbienteDev = "C:\Cerrado Tecnologia\Petromovel\PetromovelAuto\bin\Release\PetromovelAuto.exe"
    xCaminho = "C:\Cerrado Tecnologia\Petromovel\PetromovelAuto\PetromovelAuto.exe"
    
    xSaiLoop = False
    Do Until xSaiLoop = True
        If gArqTxt.FileExists(xCaminho) Then
            xRetVal = Shell(xCaminho, vbMinimizedNoFocus)
        ElseIf gArqTxt.FileExists(xCaminhoAmbienteDev) Then
            xRetVal = Shell(xCaminhoAmbienteDev, vbMinimizedNoFocus)
        Else
            MsgBox "PetromovelAuto não encontrado", vbCritical, "PetromovelAuto não configurado"
            xSaiLoop = True
        End If
        
        If pExecutarComunicacaoBanco Then
            Call AguardaMS(2000)
            If ComunicaPetromovelAutoBD = True Then
                Call GravaAuditoria(1, Me.name, 26, "PetromovelAuto Aberto e Respondendo normalmente")
                LoopAbrePetromovelAuto = True
                xSaiLoop = True
            Else
                'If (MsgBox("Deseja tentar abrir programa o PetromovelAuto novamente?", vbQuestion + vbYesNo + vbDefaultButton1, "Erro ao Abrir Programa PetromovelAuto!") = vbNo) Then
                    'Call GravaAuditoria(1, Me.name, 26, "Usuário desistiu de abrir PetromovelAuto")
                    Call GravaAuditoria(1, Me.name, 26, "Não foi possível ativar PetromovelAuto")
                    xSaiLoop = True
                'End If
            End If
        Else
            xSaiLoop = True
            LoopAbrePetromovelAuto = True
        End If
    Loop
End Function
Private Sub LimpaMSFlexGrid()
    Dim i As Integer
    MSFlexGrid.WordWrap = True
    MSFlexGrid.Rows = 2
    MSFlexGrid.Row = 1
    For i = 0 To 4
        MSFlexGrid.Col = i
        MSFlexGrid.Text = ""
    Next
    MSFlexGrid.RowHeight(0) = 650
    MSFlexGrid.Row = 0
    i = 0
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Bico"
    MSFlexGrid.ColWidth(i) = 400
    MSFlexGrid.ColAlignment(i) = 4
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Produto"
    MSFlexGrid.ColWidth(i) = 2150
    MSFlexGrid.ColAlignment(i) = 1
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Valor"
    MSFlexGrid.ColWidth(i) = 1000
    MSFlexGrid.ColAlignment(i) = 7
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Data"
    MSFlexGrid.ColWidth(i) = 1000
    MSFlexGrid.ColAlignment(i) = 4
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Hora"
    MSFlexGrid.ColWidth(i) = 800
    MSFlexGrid.ColAlignment(i) = 4
    MSFlexGrid.Row = 1
    MSFlexGrid.Col = 0
End Sub
Private Sub LimpaTela()
    cboTipoSubEstoque.ListIndex = -1
    txt_cliente.Text = ""
    dtcboCliente.BoundText = ""
    txt_produto.Text = ""
    dtcboProduto.BoundText = ""
    txt_valor_unitario.Text = ""
    txt_quantidade.Text = ""
    txt_valor_total.Text = ""
    txt_valor_desconto.Text = ""
    lPlacaLetra = ""
    lPlacaNumero = 0
    lKMVeiculo = 0
End Sub
Private Sub MarcaCelulaFlexGrid()
    If MSFlexGrid.TextMatrix(MSFlexGrid.Row, 0) <> "" Then
        If MovimentoAbastecimento.LocalizarCodigo(g_empresa, CDate(MSFlexGrid.TextMatrix(MSFlexGrid.Row, 3)), CDate(MSFlexGrid.TextMatrix(MSFlexGrid.Row, 4)), Val(MSFlexGrid.TextMatrix(MSFlexGrid.Row, 0))) Then
            lbl_fila_valor.Caption = Format(MovimentoAbastecimento.ValorUnitario, "###,##0.000")
            lbl_fila_litros.Caption = Format(MovimentoAbastecimento.Quantidade, "###,##0.000")
            lbl_fila_total.Caption = Format(MovimentoAbastecimento.ValorTotal - MovimentoAbastecimento.ValorDesconto, "###,##0.00")
        Else
            MsgBox "Não foi possível localizar abastecimento", vbInformation, "Erro de Integridade!"
        End If
    End If
End Sub
Private Sub MontaCupomVideo(x_numero_cupom As Long, x_data As Date)
    Dim i As Integer
    Dim i2 As Integer
    Dim x_string As String
    Dim x_string2 As String
    Dim xDescontoCupom As Currency
    Dim xOrdem As Integer
    
    i = 0
    lTotalCupom = 0
    xDescontoCupom = 0
    txt_cupom_fiscal.Text = ""
    xOrdem = 0
    
    Do Until MovCupomFiscal.LocalizarNumeroProximaOrdem(g_empresa, lNumeroPDV, x_numero_cupom, x_data, xOrdem) = False
        i = i + 1
        If i = 1 Then
            x_string = "                    N F C e          "
            txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + x_string + Chr(13) + Chr(10)
            txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + Chr(13) + Chr(10)
            x_string = "Data: " + Format(MovCupomFiscal.Data, "dd/mm/yyyy") + "        Hora: " + Format(MovCupomFiscal.Hora, "hh:mm:ss")
            txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + x_string + Chr(13) + Chr(10)
            x_string = "Número da NFCe: " + Format(MovCupomFiscal.NumeroCupom, "###,000")
            txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + x_string + Chr(13) + Chr(10)
            txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + "------------------------------------------------" + Chr(13) + Chr(10)
            txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + "ITEM   CÓDIGO             DESCRIÇÃO             " + Chr(13) + Chr(10)
            txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + "      QTDxUNITÁRIO       ST          VALOR( R$) " + Chr(13) + Chr(10)
            txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + "------------------------------------------------" + Chr(13) + Chr(10)
        End If
        x_string = Space(48)
        Mid(x_string, 1, 3) = Format(MovCupomFiscal.Ordem, "000")
        Mid(x_string, 5, 4) = Format(MovCupomFiscal.CodigoProduto, "###0")
        If Produto.LocalizarCodigo(MovCupomFiscal.CodigoProduto) Then
            Mid(x_string, 10, 40) = Produto.Nome
        End If
        txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + x_string + Chr(13) + Chr(10)
        x_string = Space(48)
        x_string2 = Format(MovCupomFiscal.Quantidade, "00000.0")
        If Mid(x_string2, 7, 1) = 0 Then
            i2 = Len(Format(MovCupomFiscal.Quantidade, "######0"))
            Mid(x_string, 1 + 7 - i2, i2) = Format(MovCupomFiscal.Quantidade, "######0")
        Else
            i2 = Len(Format(MovCupomFiscal.Quantidade, "####0.0"))
            If i2 > 7 Then 'casa do milhao (normalmente erro ao gerar encerrante
                            'que resulda em um cupom fiscal muito alto
                Mid(x_string, 1, 7) = "*ERRO* "
            Else
                Mid(x_string, 1 + 7 - i2, i2) = Format(MovCupomFiscal.Quantidade, "####0.0")
            End If
        End If
        Mid(x_string, 8, 3) = Mid(Produto.Unidade, 1, 2) + "x"
        x_string2 = Format(MovCupomFiscal.ValorUnitario, "00000000000.000")
        If Mid(x_string2, 15, 1) = 0 Then
            Mid(x_string, 11, 15) = Format(MovCupomFiscal.ValorUnitario, "###########0.00")
        Else
            Mid(x_string, 11, 15) = Format(MovCupomFiscal.ValorUnitario, "##########0.000")
        End If
        If Aliquota.LocalizarCodigo(lSerieECF, MovCupomFiscal.CodigoAliquota) Then
            Mid(x_string, 26, 2) = Aliquota.CodigoFiscal
        End If
        i2 = Len(Format(MovCupomFiscal.ValorTotal, "###########0.00"))
        Mid(x_string, 33 + 15 - i2, i2) = Format(MovCupomFiscal.ValorTotal, "###########0.00")
        txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + x_string + Chr(13) + Chr(10)
        If MovCupomFiscal.ItemCancelado Then
            x_string = Space(48)
            Mid(x_string, 1, 15) = "CANCELADO ITEM:"
            Mid(x_string, 16, 3) = Format(MovCupomFiscal.Ordem, "000")
            i2 = Len(Format(-MovCupomFiscal.ValorTotal, "###########0.00"))
            Mid(x_string, 32 + 16 - i2, i2) = Format(-MovCupomFiscal.ValorTotal, "###########0.00")
            txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + x_string + Chr(13) + Chr(10)
            lTotalCupom = lTotalCupom - MovCupomFiscal.ValorTotal
        End If
        lTotalCupom = lTotalCupom + MovCupomFiscal.ValorTotal
        xDescontoCupom = xDescontoCupom + MovCupomFiscal.ValorDesconto
        xOrdem = xOrdem + 1
    Loop
    If i > 0 Then
        ''txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + "------------------------------------------------" + Chr(13) + Chr(10)
        ''x_string = Space(48)
        ''Mid(x_string, 1, 15) = "T O T A L    R$"
        ''i2 = Len(Format(lTotalCupom, "###########0.00"))
        ''Mid(x_string, 33 + 15 - i2, i2) = Format(lTotalCupom, "###########0.00")
        ''txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + x_string + Chr(13) + Chr(10)
        If xDescontoCupom = 0 Then
            txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + "------------------------------------------------" + Chr(13) + Chr(10)
            x_string = Space(48)
            Mid(x_string, 1, 16) = "T O T A L     R$"
            i2 = Len(Format(lTotalCupom, "###########0.00"))
            Mid(x_string, 33 + 15 - i2, i2) = Format(lTotalCupom, "###########0.00")
            txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + x_string + Chr(13) + Chr(10)
        Else
            txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + "------------------------------------------------" + Chr(13) + Chr(10)
            x_string = Space(48)
            Mid(x_string, 1, 16) = "TOTAL BRUTO   R$"
            i2 = Len(Format(lTotalCupom, "###########0.00"))
            Mid(x_string, 33 + 15 - i2, i2) = Format(lTotalCupom, "###########0.00")
            txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + x_string + Chr(13) + Chr(10)
            If xDescontoCupom > 0 Then
                x_string = Space(48)
                Mid(x_string, 1, 16) = "DESCONTO      R$"
                i2 = Len(Format(xDescontoCupom, "###########0.00"))
                Mid(x_string, 33 + 15 - i2, i2) = Format(xDescontoCupom, "###########0.00")
                txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + x_string + Chr(13) + Chr(10)
            End If
            'txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + "------------------------------------------------" + Chr(13) + Chr(10)
            x_string = Space(48)
            Mid(x_string, 1, 16) = "TOTAL LIQUIDO R$"
            i2 = Len(Format(lTotalCupom - xDescontoCupom, "###########0.00"))
            Mid(x_string, 33 + 15 - i2, i2) = Format(lTotalCupom - xDescontoCupom, "###########0.00")
            txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + x_string + Chr(13) + Chr(10)
        End If
        If Val(MovCupomFiscal.FormaPagamento) > 0 Then
            x_string = Space(48)
            If MovCupomFiscal.FormaPagamento = 1 Then
                Mid(x_string, 1, 21) = "Dinheiro             "
            ElseIf MovCupomFiscal.FormaPagamento = 2 Then
                Mid(x_string, 1, 21) = "Cheque à Vista       "
            ElseIf MovCupomFiscal.FormaPagamento = 3 Then
                Mid(x_string, 1, 21) = "Cheque Pré-Datado    "
            ElseIf MovCupomFiscal.FormaPagamento = 4 Then
                Mid(x_string, 1, 21) = "Cartão de Crédito    "
            ElseIf MovCupomFiscal.FormaPagamento = 5 Then
                Mid(x_string, 1, 21) = "Nota de Abastecimento"
            End If
            i2 = Len(Format(MovCupomFiscal.ValorRecebido, "###########0.00"))
            Mid(x_string, 33 + 15 - i2, i2) = Format(MovCupomFiscal.ValorRecebido, "###########0.00")
            txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + x_string + Chr(13) + Chr(10)
            If MovCupomFiscal.FormaPagamento = 2 Or MovCupomFiscal.FormaPagamento = 3 Then
                x_string = Space(48)
                Mid(x_string, 1, 14) = "Cheque Número:"
                Mid(x_string, 15, 6) = MovCupomFiscal.NumeroCheque
                Mid(x_string, 23, 12) = "-  Telefone:"
                Mid(x_string, 35, 14) = fMascaraTelefone(MovCupomFiscal.Telefone)
                txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + x_string + Chr(13) + Chr(10)
            End If
            x_string = Space(48)
            Mid(x_string, 1, 21) = "Valor Recebido  R$   "
            i2 = Len(Format(MovCupomFiscal.ValorRecebido, "###########0.00"))
            Mid(x_string, 33 + 15 - i2, i2) = Format(MovCupomFiscal.ValorRecebido, "###########0.00")
            txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + x_string + Chr(13) + Chr(10)
            x_string = Space(48)
            Mid(x_string, 1, 21) = "Troco  R$            "
            ''i2 = Len(Format(![Valor Recebido] - lTotalCupom, "###########0.00"))
            ''Mid(x_string, 33 + 15 - i2, i2) = Format(![Valor Recebido] - lTotalCupom, "###########0.00")
            ''txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + x_string + Chr(13) + Chr(10)
            i2 = Len(Format(MovCupomFiscal.ValorRecebido + xDescontoCupom - lTotalCupom, "###########0.00"))
            Mid(x_string, 33 + 15 - i2, i2) = Format(MovCupomFiscal.ValorRecebido + xDescontoCupom - lTotalCupom, "###########0.00")
            txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + x_string + Chr(13) + Chr(10)
        End If
        txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + "------------------------------------------------" + Chr(13) + Chr(10)
        'txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + "Cerrado Informática - (062) 8436-4444           " + Chr(13) + Chr(10)
        x_string = ""
        x_string = ReadINI("REVENDA", "Nome", gArquivoIni)
        If Len(x_string) > 0 Then
            x_string2 = x_string
            x_string = ""
            x_string = ReadINI("REVENDA", "Telefone", gArquivoIni)
            If Len(x_string) > 0 Then
                x_string2 = x_string2 & " - " & x_string
            End If
        Else
            x_string2 = "Cerrado Tecnologia - (62)98112-5453             "
        End If
        txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + x_string2 + Chr(13) + Chr(10)
    End If
End Sub
Private Sub MostraTelaSenha()
    Call AbilitaMenu(False)
    frm_ponto.Enabled = True
    frm_ponto.ZOrder 0
    txt_funcionario_ponto = ""
    dtcboFuncionario.BoundText = 0
    txt_senha_ponto.Text = ""
    mnuPontoFuncionario.Enabled = False
    frmDados.Enabled = False
    frm_fechamento_cupom.Enabled = False
    txt_cupom_fiscal.Enabled = False
    txt_funcionario_ponto.SetFocus

    cmdIniciaProcessoFinalizacaoNFCe.Enabled = False
    btnAbortaNFCe.Enabled = False
    
    If lDescontoPostoAki Then
        Call PreencheCboFormaPagamento
    End If
End Sub
Private Sub MontaDadosTCS(ByVal pNumeroCupom As Long, ByVal pData As Date)
    Dim xOrdem As Integer
    Dim xString As String
    
    xString = Space(6)
    'Call CriaLogCupom("Bematech_FI_NumeroOperacoesNaoFiscais")
    'BemaRetorno = Bematech_FI_NumeroOperacoesNaoFiscais(xString)
    'Call CriaLogCupom("Bematech_FI_NumeroOperacoesNaoFiscais(xString) - xString=" & xString & " - BemaRetorno=" & BemaRetorno)
    'If BemaRetorno <> 1 Then
        xString = "000000"
    'End If
    lContadorNaoFiscal = xString
    
    xOrdem = 0
    lDadosTCS = ""
    Do Until MovCupomFiscal.LocalizarNumeroProximaOrdem(g_empresa, lNumeroPDV, pNumeroCupom, pData, xOrdem) = False
        lDadosTCS = lDadosTCS & MovCupomFiscal.CodigoProduto & "|@|"
        lDadosTCS = lDadosTCS & Format(MovCupomFiscal.Quantidade, "000000.00") * 100 & "|@|"
        If TicketCarDePara.LocalizarCodigo(MovCupomFiscal.CodigoProduto) Then
            lDadosTCS = lDadosTCS & Format(TicketCarDePara.CodigoTCS, "0000") & "|@|"
        End If
        If Produto.LocalizarCodigo(MovCupomFiscal.CodigoProduto) Then
            lDadosTCS = lDadosTCS & Produto.Nome & "|@|"
        End If
        lDadosTCS = lDadosTCS & Format(MovCupomFiscal.ValorTotal, "000000.00") * 100 & "|@|"
        If Produto.Unidade = "SRV" Then
            lDadosTCS = lDadosTCS & "1" & "|@|" '0=Produto, 1=Serviço
        Else
            lDadosTCS = lDadosTCS & "0" & "|@|" '0=Produto, 1=Serviço
        End If
        lDadosTCS = lDadosTCS & Format(MovCupomFiscal.ValorUnitario, "000000.000") * 1000 & "|@@|"
        xOrdem = xOrdem + 1
    Loop
End Sub
Private Sub MostraMensagensAguarde(ByVal pTitulo As String, ByVal pMensagem As String, ByVal pSuperior As Currency, ByVal pEsquerda As Currency, ByVal pLargura As Currency, ByVal pAltura As Currency)
    frameAguarde.Top = pSuperior
    frameAguarde.Left = pEsquerda
    frameAguarde.Width = pLargura
    frameAguarde.Height = pAltura
    
    DoEvents
    
    lblTituloAguarde.Width = pLargura - 200
    lblTituloAguarde.Top = pAltura * 0.25
    lblMensagemAguarde.Width = pLargura - 300
    lblMensagemAguarde.Top = pAltura * 0.6
    DoEvents
    lblTituloAguarde.Caption = pTitulo
    lblMensagemAguarde.Caption = pMensagem
    DoEvents
End Sub
Private Sub MudaHorarioVeraoAutomatico()
    Dim xMensagem As String
    Dim xHoraInicial As Date
    Dim xSaiLoop As Boolean
    Dim xMensagemCupomAnterior As String
                    
    lExisteMudancaHorarioVerao = False
    TimerAutomacao.Interval = 0
    TimerAutomacao.Enabled = False
    xMensagem = "ECF programado para entrar no horário de verão!"
    If MovHorarioVerao.EntradaHorarioVerao Then
        xMensagem = "ECF programado para entrar no"
    Else
        xMensagem = "ECF programado para sair do"
    End If
    xMensagem = xMensagem & " horário de verão!" & vbCrLf
    xMensagem = xMensagem & "Na data " & Format(MovHorarioVerao.DataParaMudancaHorario, "dd/MM/yyyy") & " às " & Format(MovHorarioVerao.HoraParaMudancaHorario, "HH:mm:ss") & vbCrLf & vbCrLf
    xMensagem = xMensagem & "Favor não utilizar o sistema neste computador até a mudança do horário." & vbCrLf
    xMensagem = xMensagem & "Após a conclusão deste processo irá aparecer uma mensagem autorizando o uso deste computador." & vbCrLf
    xMensagem = xMensagem & "Por este motivo não será permitido tirar qualquer tipo documento neste ECF."
    xMensagemCupomAnterior = txt_cupom_fiscal.Text
    txt_cupom_fiscal.ZOrder 0
    txt_cupom_fiscal.Enabled = True
    txt_cupom_fiscal.Top = 10
    txt_cupom_fiscal.Left = 10
    txt_cupom_fiscal.Height = Me.Height - 1000
    txt_cupom_fiscal.Width = Me.Width - 200
    txt_cupom_fiscal.Text = xMensagem
    
    
    'Aguarda 3 segundos
    xHoraInicial = Time
    Do Until DateDiff("s", xHoraInicial, Time) >= 3
        DoEvents
    Loop
    
    'Verifica se é pra imprimir Redução Z
'    If MovHorarioVerao.DataParaImpressaoReducaoZ <> CDate("00:00:00") And MovHorarioVerao.ComandoReducaoZConcluido = False Then
'        'Aguarda até chegar no horário programado para Imprimir Redução Z
'        xSaiLoop = False
'        Do Until xSaiLoop = True
'            DoEvents
'            If Date >= MovHorarioVerao.DataParaImpressaoReducaoZ Then
'                If Format(Time, "HH:mm:ss") >= Format(MovHorarioVerao.HoraParaImpressaoReducaoZ, "HH:mm:ss") Then
'                    'MsgBox "executa comando para iprimir a redução Z"
'                    'ImprimeReducaoZ
'                    MovHorarioVerao.ComandoReducaoZConcluido = True
'                    If Not MovHorarioVerao.Alterar(g_empresa, lNumeroPDV, MovHorarioVerao.DataParaInicioBloqueio, MovHorarioVerao.HoraParaInicioBloqueio) Then
'                        MsgBox "Erro ao concluir mudança de impressão da redução Z.", vbCritical, "Erro de Integridade!"
'                    End If
'                    xSaiLoop = True
'                    'Aguarda 5 segundos
'                    xHoraInicial = Time
'                    Do Until DateDiff("s", xHoraInicial, Time) >= 5
'                        DoEvents
'                    Loop
'                End If
'            End If
'            DoEvents
'        Loop
'    End If
    
    'Aguarda até chegar no horário programado para Mudar Horário de Verão
    xSaiLoop = False
    Do Until xSaiLoop = True
        DoEvents
        If Date >= MovHorarioVerao.DataParaMudancaHorario Then
            If Format(Time, "HH:mm:ss") >= Format(MovHorarioVerao.HoraParaMudancaHorario, "HH:mm:ss") Then
                'MsgBox "executa comando para mudança de horario de verao"
                MudaHorarioVeraoECF
                MovHorarioVerao.ComandoVeraoConcluido = True
                If Not MovHorarioVerao.Alterar(g_empresa, lNumeroPDV, MovHorarioVerao.DataParaInicioBloqueio, MovHorarioVerao.HoraParaInicioBloqueio) Then
                    MsgBox "Erro ao concluir mudança de horário de verão.", vbCritical, "Erro de Integridade!"
                End If
                xSaiLoop = True
            End If
        End If
        DoEvents
    Loop
    
    txt_cupom_fiscal.Text = "Mudança de horário de verão concluída com sucesso!"
    MsgBox "Mudança de horário de verão concluída com sucesso!" & vbCrLf & "Computador liberado para uso.", vbOKOnly + vbInformation, "Programação Automática de Verão Concluído"
    txt_cupom_fiscal.Top = 60
    txt_cupom_fiscal.Left = 5955
    txt_cupom_fiscal.Height = 6915
    txt_cupom_fiscal.Width = 5835
    txt_cupom_fiscal.Text = xMensagemCupomAnterior
    lbl_mensagem.ToolTipText = ""
    
    TimerAutomacao.Interval = 1000
    TimerAutomacao.Enabled = True
    mnuSenha_Click
    'txt_funcionario_ponto.SetFocus
End Sub
Private Sub MudaHorarioVeraoECF()
    Dim xRetorno As Long
    Dim HorarioVerao As Byte
    Dim x_dia, x_mes, x_ano, x_hora, x_minuto, x_segundo As Integer
    
    Call GravaAuditoria(1, Me.name, 26, " Horário de Verão. Func.:" & l_nome_funcionario)
    If lImpBematech Then
        'Call CriaLogCupom("Bematech_FI_ProgramaHorarioVerao")
        'BemaRetorno = Bematech_FI_ProgramaHorarioVerao
        'Call CriaLogCupom("Bematech_FI_ProgramaHorarioVerao - BemaRetorno=" & BemaRetorno)
'    ElseIf lImpSchalter Then
'        x_hora = Format(Mid(msk_hora, 1, 2) + 1, "00")
'        x_minuto = Mid(msk_hora, 4, 2)
'        x_segundo = Mid(msk_hora, 7, 2)
'        x_dia = Format(Format(lDataCupom, "dd"), "00")
'        x_mes = Format(Format(lDataCupom, "mm"), "00")
'        x_ano = Format(lDataCupom, "yyyy")
'        xRetorno = ecfAcertaData(x_dia, x_mes, x_ano, x_hora, x_minuto, x_segundo)
'    ElseIf lImpMecaf Then
'        HorarioVerao = Asc("+")
'        xRetorno = ProgramaHorarioVerao(HorarioVerao)
    ElseIf lImpQuick Then
        If Not EcfQuickAcertaHorarioVerao Then
            MsgBox "Não foi possível mudar o horário de/para verão!", vbCritical, "Comando não Executado!"
        End If
'    ElseIf lImpElgin Then
'        BemaRetorno = Elgin_ProgramaHorarioVerao
'        If BemaRetorno <> 1 Then
'            MsgBox "Não foi possível mudar o horário de/para verão!", vbCritical, "Comando não Executado!"
'        End If
    End If
End Sub
Private Sub NovoCupom()
    Dim i As Integer
    
    
    
    lData = Format(Now, "dd/mm/yyyy") 'ALEX - NFCe 'Space(6)
    lHora = Format(Now, "HH:mm:ss")  'ALEX - NFCe 'Space(6)
    lDataCupom = lData
    
    'ImprimeProgramaFormaPagamento
    LimpaTela
'    If BuscaNumeroCupom = "ECF SEM COMUNICACAO" Then
'        MsgBox "Não foi possível comunicar com a Impressora Fiscal!", vbApplicationModal, "ECF sem comunicação!"
'        mnuSenha_Click
'        Exit Sub
'    End If

    If lOrdem = 0 Then
        For i = 0 To gQUANTIDADE_MAXIMA_ITENS_NFCE
            lTotalItem(i) = 0
        Next
    End If
    
    'Aqui chamo ExisteCupom, pra atribuir valores a combo e outroas variaveis
    'Pois no segundo ítem, esses valores foram zerados
    ExisteCupom
    If lOrdem > 0 Then
        txt_produto.SetFocus
    Else
        If cboTipoSubEstoque.ListCount > 1 And UCase(Funcionario.Cargo) Like "*TROCADOR*" Then
            cboTipoSubEstoque.ListIndex = 1
        Else
            cboTipoSubEstoque.ListIndex = 0
        End If
        txt_cliente.Text = ""
        txt_cliente.SetFocus
    End If
    BuscaPeriodo
'    If lNumeroCupom = 0 Then
'        CancelaCupom
'    Else
'        'If CDate(lHora) >= CDate("00:00:00") And CDate(lHora) < CDate("14:00:00") Then
'        '    lPeriodo = 1
'        'ElseIf CDate(lHora) >= CDate("14:00:00") And CDate(lHora) < CDate("23:00:00") Then
'        '    lPeriodo = 2
'        'End If
''        If CDate(lHora) >= CDate("00:00:00") And CDate(lHora) < CDate("06:00:00") Then
''            lPeriodo = 1
''        ElseIf CDate(lHora) >= CDate("06:00:00") And CDate(lHora) < CDate("14:00:00") Then
''            lPeriodo = 2
''        ElseIf CDate(lHora) >= CDate("14:00:00") And CDate(lHora) < CDate("22:00:00") Then
''            lPeriodo = 3
''        Else
''            lPeriodo = 4
''            'If Format(lHora, "hh") >= 0 And Format(lHora, "hh") < 6 Then
''            '    lData = lData - 1
''            'End If
''        End If
'    End If
    Me.Caption = "Nota Fiscal do Consumidor Eletrônica - " & l_nome_funcionario & " | Caixa: " & Val(g_cfg_periodo_i) & " Em: " & Format(g_cfg_data_i, "dd/mm/yyyy")
End Sub
Private Sub GravaItem()
    Dim xGrava As Boolean
    Dim xLocalErro As Integer


    On Error GoTo FileError
    
    xGrava = False
'    'Teste evita emitir um cupom em cima de um documento existente.
'    'Ex. Apos sair uma reducao Z automatica, gerar um cupom
'    'no banco de dados com o mesmo numero da reducao Z
'    If lOrdem = 1 Then
'        If DateDiff("n", lHoraPegouNumeroCupom, Now) > 0 Then
'            BuscaNumeroCupom
'        End If
'    End If
    xLocalErro = 1
    If ValidaCampos Then
        xLocalErro = 5
        If VerificaLiberacaoDigitacao2 Then
            xLocalErro = 10
            BuscaNumeroCupom
            xLocalErro = 20
            If lAutomacaoFlagVendaAutomatica = True Then
                xLocalErro = 25
                AtualTabe
                xLocalErro = 30
                MovCupomFiscal.Nome = Format(lAutomacaoDataEmAcerto, "dd/mm/yyyy") & "-" & Format(lAutomacaoHoraEmAcerto, "HH:mm:ss") & "-" & lAutomacaoBicoEmAcerto & "-" & lAutomacaoTempoAbastecimentoEmAcerto & "-" & lAutomacaoFormaImpressao
                xLocalErro = 31
                'NFCE - ALEX
                If MovimentoAbastecimento.LocalizarCodigo(g_empresa, lAutomacaoDataEmAcerto, lAutomacaoHoraEmAcerto, lAutomacaoBicoEmAcerto) Then
                    xLocalErro = 32
                    MovCupomFiscal.Telefone = Format(MovimentoAbastecimento.Encerrante, "#######.00")
                End If
                
                xLocalErro = 33
                If Bomba.LocalizarCodigo(g_empresa, lAutomacaoBicoEmAcerto) Then
                    xLocalErro = 34
                    MovCupomFiscal.NumeroCheque = Format(Bomba.CodigoFisicoBomba, "00") & Format(Bomba.Codigo, "00") & Format(Bomba.NumeroTanque, "00")
                End If
                
                
                xLocalErro = 100
                If ImprimeCupomFiscal Then
                '25/06/14^^
                    xLocalErro = 110
                    If MovCupomFiscal.Incluir Then
                        If Produto.CodigoGrupo = lGrupoCombustivel Then
                            Call IncluiSaidaBomba(MovCupomFiscal.CodigoProduto, MovCupomFiscal.Quantidade)
                        End If
                        xLocalErro = 130
                        If GravaItemCupom Then
                        End If
                        
                        xLocalErro = 140
                        If lOrdem <= 1 Then 'ALEX - DOCELETRONICO
                            xLocalErro = 150
                            PreencheDadosDocumentoEletronicoCabecalho
                            xLocalErro = 160
                            If MovDocEletronicoCabecalho.Incluir Then
                               btnAbortaNFCe.Enabled = True 'ALEX - ABORTAR
                               mnuSenha.Enabled = False
                               Call GravaDocumentoEletronicoEvento(MovDocEletronicoCabecalho, EVENTO_NFCE.ABERTA)
                                If Not GravaDocumentoEletronicoItem Then
                                    MsgBox "Não foi possível incluir o item " & lOrdem & ".", vbInformation, "Erro de Integridade."
                                End If
                            Else
                                MsgBox "Não foi possível incluir dados do cabeçalho NFCE.", vbInformation, "Erro de Comunicação."
                            End If
                        Else
                            If Not GravaDocumentoEletronicoItem Then
                               MsgBox "Não foi possível incluir o item " & lOrdem & ".", vbInformation, "Erro de Integridade."
                            End If
                        End If
                        
                        xLocalErro = 200
                        If lBaixaAutomaticaNoEstoque = True Then
                            Call SubtraiEstoque(CLng(txt_produto.Text), fValidaValor(txt_quantidade.Text), cboTipoSubEstoque.ItemData(cboTipoSubEstoque.ListIndex))
                        End If
                        Call BuscaRegistro(lNumeroCupom, lData, lOrdem)
                        xLocalErro = 220
                        NovoCupom
                        Call MontaCupomVideo(lNumeroCupom, lData)
                    Else
                        MsgBox "Não foi possível incluir o cupom fiscal.", vbInformation, "Erro de Integridade."
                        NovoCupom
                        Call MontaCupomVideo(lNumeroCupom, lData)
                    End If
                    xLocalErro = 300
                
                
                
                Else
                    xLocalErro = 350
                    MsgBox "Não foi possível imprimir este ítem do cupom fiscal.", vbInformation, "Erro de Comunicação!"
                    NovoCupom
                    xLocalErro = 360
                    Call MontaCupomVideo(lNumeroCupom, lData)
                End If
                
                
            Else
                xLocalErro = 500
                'NFCE - ALEX
                Dim EncerranteTeste As Currency
                EncerranteTeste = 1000.111
                MovCupomFiscal.Telefone = Format(EncerranteTeste, "#######.00")    'Adicionado para teste em ambiente de desenvolvimento
                MovCupomFiscal.NumeroCheque = "010101" 'Adicionado para teste em ambiente de desenvolvimento
                
                xLocalErro = 520
                If lCodigoBarra Then
                    xGrava = True
                Else
                    If (MsgBox("Deseja gravar este ítem?", vbYesNo + vbDefaultButton1 + vbQuestion, "Grava Ítem do NFCe")) = vbYes Then
                        xGrava = True
                    End If
                End If
                xLocalErro = 550
                If xGrava Then
                    xLocalErro = 560
                    'Imprime Vale Abastecimento
                    If UCase(Trim(dtcboProduto.Text)) = "VALE ABASTECIMENTO" Then
                        Call ImpValeTroco
                        mnuSenha_Click
                        Exit Sub
                    End If
                    xLocalErro = 580
                    AtualTabe
                    xLocalErro = 585
                    If ImprimeCupomFiscal Then
                        xLocalErro = 600
                        If MovCupomFiscal.Incluir Then
                            xLocalErro = 605
                            If Produto.CodigoGrupo = lGrupoCombustivel Then
                                Call IncluiSaidaBomba(MovCupomFiscal.CodigoProduto, MovCupomFiscal.Quantidade)
                            End If
                            If GravaItemCupom Then
                            End If
                            
                            xLocalErro = 650
                            If lOrdem <= 1 Then
                                PreencheDadosDocumentoEletronicoCabecalho
'                                If Not MovDocEletronicoCabecalho.Incluir Then
'                                     Call CriaLogSGP("[GravaItem]", "Não foi possível incluir dados do cabeçalho NFCE", "")
'                                     MsgBox "Não foi possível incluir dados do cabeçalho NFCE.", vbInformation, "Erro de Integridade."
'                                End If
                                If MovDocEletronicoCabecalho.Incluir Then
                                   btnAbortaNFCe.Enabled = True 'ALEX - ABORTAR
                                   mnuSenha.Enabled = False
                                   Call GravaDocumentoEletronicoEvento(MovDocEletronicoCabecalho, EVENTO_NFCE.ABERTA)
                                   If Not GravaDocumentoEletronicoItem Then
                                      Call CriaLogSGP("[GravaItem]", "Não foi possível incluir o item " & lOrdem & ".", "")
                                      MsgBox "Não foi possível incluir o item " & lOrdem & ".", vbInformation, "Erro de Integridade."
                                   End If
                                Else
                                    Call CriaLogSGP("[GravaItem]", "Não foi possível incluir dados do cabeçalho NFCE", "")
                                    MsgBox "Não foi possível incluir dados do cabeçalho NFCE.", vbInformation, "Erro de Comunicação."
                                End If
                            Else
                                If Not GravaDocumentoEletronicoItem Then
                                    Call CriaLogSGP("[GravaItem]", "Não foi possível incluir o item " & lOrdem & ".", "")
                                    MsgBox "Não foi possível incluir o item " & lOrdem & ".", vbInformation, "Erro de Integridade."
                                End If
                            End If
    
                            xLocalErro = 680
                            If lBaixaAutomaticaNoEstoque = True Then
                                Call SubtraiEstoque(CLng(txt_produto.Text), fValidaValor(txt_quantidade.Text), cboTipoSubEstoque.ItemData(cboTipoSubEstoque.ListIndex))
                            End If
                            Call BuscaRegistro(lNumeroCupom, lData, lOrdem)
                            If Produto.CodigoGrupo = lGrupoCombustivel Then
                            Else
                                If lBaixaAutomaticaNoEstoque = True Then
                                    Call AtualizaTabelaVendaProduto
                                End If
                            End If
                            xLocalErro = 700
                            NovoCupom
                            Call MontaCupomVideo(lNumeroCupom, lData)
                        Else
                            xLocalErro = 750
                            MsgBox "Não foi possível incluir o cupom fiscal.", vbInformation, "Erro de Integridade."
                            NovoCupom
                            Call MontaCupomVideo(lNumeroCupom, lData)
                        End If
                    Else
                        xLocalErro = 800
                        MsgBox "Não foi possível imprimir este ítem do cupom fiscal.", vbInformation, "Erro de Comunicação."
                        NovoCupom
                        Call MontaCupomVideo(lNumeroCupom, lData)
                    End If
                    
                Else
                    txt_produto.SetFocus
                End If
            End If
        End If
    End If
    lOrigemAutomacao = False
    Exit Sub
FileError:
    Call CriaLogCupom("Cupom Fiscal: ERRO ao gravar item." & " - xLocalErro: " & xLocalErro)
    MsgBox "Erro desconhecido!", vbInformation, "Erro desconhecido!"
 Exit Sub
End Sub
Private Function GravaItemCupom() As Boolean
    Dim i As Integer
    On Error GoTo FileError
    
    GravaItemCupom = False
    MovCupomFiscalItem.Empresa = g_empresa
    MovCupomFiscalItem.NumeroCupom = lNumeroCupom
    MovCupomFiscalItem.Ordem = lOrdem
    MovCupomFiscalItem.Data = lData
    MovCupomFiscalItem.CodigoProduto = CLng(txt_produto.Text)
    MovCupomFiscalItem.ValorUnitario = fValidaValor4(txt_valor_unitario.Text)
    MovCupomFiscalItem.Quantidade = fValidaValor(txt_quantidade.Text)
    MovCupomFiscalItem.ValorTotal = fValidaValor2(txt_valor_total.Text)
    MovCupomFiscalItem.ItemCancelado = False
    If lDescontoItemEmbutido = 0 And lAcrescimoItemEmbutido = 0 Then
        MovCupomFiscalItem.ValorDesconto = 0
        MovCupomFiscalItem.ValorAcrescimo = 0
        MovCupomFiscalItem.DescontoEmbutido = False
    Else
        MovCupomFiscalItem.ValorDesconto = lDescontoItemEmbutido
        MovCupomFiscalItem.ValorAcrescimo = lAcrescimoItemEmbutido
        MovCupomFiscalItem.DescontoEmbutido = True
    End If
    MovCupomFiscalItem.Periodo = lPeriodo
    MovCupomFiscalItem.TipoCombustivel = Produto.TipoCombustivel
    MovCupomFiscalItem.CodigoECF = lNumeroPDV
    MovCupomFiscalItem.CodigoAliquota = Produto.CodigoAliquota
    MovCupomFiscalItem.CodigoGrupo = Produto.CodigoGrupo
    If MovCupomFiscalItem.Incluir Then
        GravaItemCupom = True
    Else
        MsgBox "Não foi possível incluir item do cupom fiscal.", vbInformation, "Erro de Integridade!"
        i = 99999
    End If
    Exit Function

FileError:
    Dim xString As String
    xString = "Numero=" & lNumeroCupom
    xString = xString & " - Ordem=" & lOrdem
    xString = xString & " - Data=" & lData
    xString = xString & " - Produto=" & CLng(txt_produto.Text)
    xString = xString & " - Quantidade=" & txt_quantidade.Text
    xString = xString & " - ValorUnitario=" & txt_valor_unitario.Text
    xString = xString & " - ValorTotal=" & txt_valor_total.Text
    Call CriaLogCupom("ERRO: Ao gravar Item do Cupom Fiscal - " & xString)
    Exit Function
End Function
Private Function GravaDocumentoEletronicoEvento(ByVal pDocumentoEltronicoCabecalho As cMovDocEletronicoCabecalho, ByVal pDescricaoEvento As EVENTO_NFCE) As Boolean

    On Error GoTo FileError

    Dim xMovEvento As New cMovDocEletronicoEvento
    Dim xPassoProcesso As String
    
    GravaDocumentoEletronicoEvento = False
    
    xPassoProcesso = ""
    
    xPassoProcesso = "1"
    With xMovEvento
        .IdEstabelecimento = pDocumentoEltronicoCabecalho.IdEstabelecimento
        .Modelo = pDocumentoEltronicoCabecalho.Modelo
        .numero = Val(lNumeroCupom)
        .Serie = pDocumentoEltronicoCabecalho.Serie
        .DataEmissao = pDocumentoEltronicoCabecalho.DataEmissao
        .Sequencia = xMovEvento.ProximaSequencia(pDocumentoEltronicoCabecalho.IdEstabelecimento, pDocumentoEltronicoCabecalho.DataEmissao, pDocumentoEltronicoCabecalho.Modelo, pDocumentoEltronicoCabecalho.Serie, lNumeroCupom)
        .DataHora = Now
        .CodigoTipoEvento = Val(pDescricaoEvento)
        .Descricao = xMovEvento.DescricaoEnumEvento(pDescricaoEvento)
    End With
    
    xPassoProcesso = "2"
    If xMovEvento.Incluir Then
        xPassoProcesso = "3"
        pDocumentoEltronicoCabecalho.CodigoUltimoEvento = Val(pDescricaoEvento)
        pDocumentoEltronicoCabecalho.ObservacaoEvento = xMovEvento.DescricaoEnumEvento(pDescricaoEvento)
        Call pDocumentoEltronicoCabecalho.DefinirUtlimoEventoDocumento(pDocumentoEltronicoCabecalho.IdEstabelecimento, pDocumentoEltronicoCabecalho.DataEmissao, pDocumentoEltronicoCabecalho.Modelo, pDocumentoEltronicoCabecalho.Serie, pDocumentoEltronicoCabecalho.numero)
        GravaDocumentoEletronicoEvento = True
    Else
        Call CriaLogECF(Date & " " & Time & " GravaDocumentoEletronicoEvento: Não foi possível incluir o evento da NFCe: " & CStr(pDescricaoEvento) & " xPassoProcesso=" & xPassoProcesso)
        MsgBox "Não foi possível incluir o evento da NFCe: " & CStr(pDescricaoEvento) & ".", vbInformation, "Erro de Integridade."
    End If
    
    Exit Function

FileError:
    Call CriaLogECF(Date & " " & Time & " GravaDocumentoEletronicoEvento: " & Error & " xPassoProcesso=" & xPassoProcesso)
    MsgBox "Não foi possível incluir o evento da NFCe: " & CStr(pDescricaoEvento) & ". ERRO= " & Err.Description, vbCritical, "Erro de Integridade."

End Function

Private Function GravaDocumentoEletronicoItem() As Boolean 'ALEX - DOCELETRONICO
    'Dim i As Integer
    Const CODIGO_FISCAL_ST As String = "FF"
    
    
    On Error GoTo FileError
    
    GravaDocumentoEletronicoItem = False
    
    MovDocEletronicoItem.IdEstabelecimento = g_empresa
    MovDocEletronicoItem.DataEmissao = lData
    MovDocEletronicoItem.Entrada = False
    MovDocEletronicoItem.Saida = True
    MovDocEletronicoItem.Modelo = MODELO_NFCE
    MovDocEletronicoItem.Serie = lSerieNFCe
    MovDocEletronicoItem.numero = lNumeroCupom
    MovDocEletronicoItem.Ordem = lOrdem
    MovDocEletronicoItem.IdClienteFornecedor = l_codigo_cliente
    MovDocEletronicoItem.MovimentacaoFisica = True
    MovDocEletronicoItem.Cfop = "5656"
    MovDocEletronicoItem.CstPis = Produto.CstPis
    MovDocEletronicoItem.CstCofins = Produto.CstCofins
    MovDocEletronicoItem.ValorTotalLiquido = fValidaValor2(txt_valor_total.Text)
    
    
'19/04/2018 - COMENTADO POIS O CORRETO SERÁ OBTER CST DO PRODUTO E A ALIQUOTA PIS E COFINS DA TABELA TRIBUTACAOPISCOFINS
    If Grupo.LocalizarCodigo(Produto.CodigoGrupo) = True Then
        MovDocEletronicoItem.Cfop = Grupo.CfopSaida

'        MovDocEletronicoItem.CSTPIS = Grupo.CstPisSaida
'        MovDocEletronicoItem.AliquotaPis = Grupo.AliquotaPis
'        MovDocEletronicoItem.CSTCOFINS = Grupo.CstCofinsSaida
'        MovDocEletronicoItem.AliquotaCofins = Grupo.AliquotaCofins
'    Else
'        MovDocEletronicoItem.CSTPIS = Produto.CSTPIS
'        MovDocEletronicoItem.AliquotaPis = 0 'OBTER ALIQUOTA
'        MovDocEletronicoItem.CSTCOFINS = Produto.CSTCOFINS
'        MovDocEletronicoItem.AliquotaCofins = 0
    End If
    
    
    If lDescontoItemEmbutido = 0 And lAcrescimoItemEmbutido = 0 Then
       MovDocEletronicoItem.ValorDesconto = 0
    Else
       MovDocEletronicoItem.ValorDesconto = 0
    End If

    
    'ICMS
    MovDocEletronicoItem.CSTICMS = Produto.CSTICMS
    MovDocEletronicoItem.CstIcmsSt = Produto.CSTICMS 'VERIFICAR SE REALEMTE É ESSE
    If Aliquota.CodigoFiscal = CODIGO_FISCAL_ST Then
        MovDocEletronicoItem.ValorBCICMS = 0
        MovDocEletronicoItem.AliquotaICMS = 0
        MovDocEletronicoItem.ValorBCICMSST = 0
        MovDocEletronicoItem.AliquotaIcmsSt = 0 'OBTER ALIQUOTA
        MovDocEletronicoItem.ValorICMSST = 0
    Else
        MovDocEletronicoItem.AliquotaICMS = Aliquota.Aliquota 'OBTER ALIQUOTA
        MovDocEletronicoItem.ValorBCICMS = IIf(Aliquota.Aliquota > 0, MovDocEletronicoItem.ValorTotalLiquido, 0) 'ALEX - BASECALCULO
        MovDocEletronicoItem.ValorBCICMSST = 0
        MovDocEletronicoItem.AliquotaIcmsSt = 0
        MovDocEletronicoItem.ValorICMSST = 0
    End If
    
    MovDocEletronicoItem.ValorICMS = RetornaValorImpostoProdutoNFCE(MovDocEletronicoItem.ValorBCICMS, MovDocEletronicoItem.AliquotaICMS) 'IIf(Aliquota.Aliquota > 0, MovDocEletronicoItem.ValorBCICMS * (Aliquota.Aliquota / 100), 0) 'VERIFICAR DE ONDE OBTER
    
    
    'PIS - 19/04/2018
    If CSTPisValidos.Item(Val(MovDocEletronicoItem.CstPis)) > 0 Then
        MovDocEletronicoItem.ValorBcPis = FormatNumber(MovDocEletronicoItem.ValorTotalLiquido, 4)
        MovDocEletronicoItem.ValorPis = FormatNumber(MovDocEletronicoItem.ValorBcPis * (CSTPisValidos.Item(Val(MovDocEletronicoItem.CstPis)) / 100), 2)
        MovDocEletronicoItem.AliquotaPis = CSTPisValidos.Item(Val(MovDocEletronicoItem.CstPis))
    Else
        MovDocEletronicoItem.ValorBcPis = 0
        MovDocEletronicoItem.ValorPis = 0
        MovDocEletronicoItem.AliquotaPis = 0
    End If

    MovDocEletronicoItem.QuantidadeBcPis = 0 'utilizado apenas para cst de pis 03.
    
    'COFINS - 19/04/2018
    If CSTCofinsValidos.Item(Val(MovDocEletronicoItem.CstCofins)) > 0 Then
        MovDocEletronicoItem.ValorBcCofins = FormatNumber(MovDocEletronicoItem.ValorTotalLiquido, 4)
        MovDocEletronicoItem.ValorCofins = FormatNumber(MovDocEletronicoItem.ValorBcCofins * (CSTCofinsValidos.Item(Val(MovDocEletronicoItem.CstCofins)) / 100), 2)
        MovDocEletronicoItem.AliquotaCofins = CSTCofinsValidos.Item(Val(MovDocEletronicoItem.CstCofins))
    Else
        MovDocEletronicoItem.ValorBcCofins = 0
        MovDocEletronicoItem.ValorCofins = 0
        MovDocEletronicoItem.AliquotaCofins = 0
    End If

    MovDocEletronicoItem.QuantidadeBcCofins = 0 'utilizado apenas para cst de pis 03.


    
    'IPI
    MovDocEletronicoItem.CSTIPI = Empty 'Produto.CSTIPI
    MovDocEletronicoItem.ValorBcIpi = 0 'fValidaValor2(txt_valor_total.Text) ''VERIFICAR SE REALEMTE É ESSE
    MovDocEletronicoItem.AliquotaIpi = 0  ''OBTER ALIQUOTA
    MovDocEletronicoItem.ValorIPI = 0 'VERIFICAR DE ONDE OBTER
    MovDocEletronicoItem.ApuracaoIpiMensal = False 'VERIFICAR SE REALEMTE É ESSE
    MovDocEletronicoItem.CodigoEnquadramentoIpi = Empty 'VERIFICAR DE ONDE OBTER
    

'19/04/2018 - COMENTADO POIS O CORRETO SERÁ OBTER CST DO PRODUTO E A ALIQUOTA PIS E COFINS DA TABELA TRIBUTACAOPISCOFINS
'    'PIS
'    If CSTPisCofinsValidos.Item(Val(MovDocEletronicoItem.CSTPIS)) = True Then
'        MovDocEletronicoItem.ValorBcPis = FormatNumber(MovDocEletronicoItem.ValorTotalLiquido, 4)
'        MovDocEletronicoItem.ValorPis = FormatNumber(MovDocEletronicoItem.ValorBcPis * (MovDocEletronicoItem.AliquotaPis / 100), 2)
'    Else
'        MovDocEletronicoItem.ValorBcPis = 0
'        MovDocEletronicoItem.ValorPis = 0
'    End If
'
'    MovDocEletronicoItem.QuantidadeBcPis = 0 'utilizado apenas para cst de pis 03.
'
'    'COFINS
'    If CSTPisCofinsValidos.Item(Val(MovDocEletronicoItem.CstCofins)) = True Then
'        MovDocEletronicoItem.ValorBcCofins = FormatNumber(MovDocEletronicoItem.ValorTotalLiquido, 4)
'        MovDocEletronicoItem.ValorCofins = FormatNumber(MovDocEletronicoItem.ValorBcCofins * (MovDocEletronicoItem.AliquotaCofins / 100), 2)
'    Else
'        MovDocEletronicoItem.ValorBcCofins = 0
'        MovDocEletronicoItem.ValorCofins = 0
'    End If
'
'    MovDocEletronicoItem.QuantidadeBcCofins = 0 'utilizado apenas para cst de cofins 03.


    
    MovDocEletronicoItem.IdProduto = CLng(txt_produto.Text)
    MovDocEletronicoItem.ValorUnitario = fValidaValor4(txt_valor_unitario.Text)
    If Produto.CodigoGrupo = 4 Then
        MovDocEletronicoItem.Quantidade = MovDocEletronicoItem.ValorTotalLiquido / MovDocEletronicoItem.ValorUnitario 'fValidaValor(txt_quantidade.Text)
    Else
        MovDocEletronicoItem.Quantidade = fValidaValor(txt_quantidade.Text)
    End If
    
    MovDocEletronicoItem.NumeroTanque = Format(1, "00")
    MovDocEletronicoItem.EncerranteFinal = Format(10000, "#######.00")
    MovDocEletronicoItem.NumeroBomba = Format(1, "00")
    MovDocEletronicoItem.NumeroBico = Format(1, "00")
    MovDocEletronicoItem.TipoCombustivel = Produto.TipoCombustivel
   
    
    If lAutomacaoFlagVendaAutomatica = True Then
        CriaLogCupom "NFCE Venda Automatica = true (BICO EM ABERTO =" & lAutomacaoBicoEmAcerto & " )" & " Empresa: " & g_empresa & " Data Automação= " & lAutomacaoDataEmAcerto & " Hora Automação: " & lAutomacaoHoraEmAcerto & " Bico Automação: " & lAutomacaoBicoEmAcerto
        
        CriaLogCupom "-- Tentar localizar ABASTECIMENTO --"
        If MovimentoAbastecimento.LocalizarCodigo(g_empresa, lAutomacaoDataEmAcerto, lAutomacaoHoraEmAcerto, lAutomacaoBicoEmAcerto) Then
           CriaLogCupom "NFCE Venda Automatica = true - ABASTECIMENTO ENCONTRADO OK"
           
           MovDocEletronicoItem.EncerranteFinal = Format(MovimentoAbastecimento.Encerrante, "#######.00")
           
           CriaLogCupom "ENCERRANTE FINAL = " & Format(MovimentoAbastecimento.Encerrante, "#######.00")
        End If
                    
        CriaLogCupom "-- Tentar localizar BOMBA -- BICO EM ABERTO = " & lAutomacaoBicoEmAcerto
        If Bomba.LocalizarCodigo(g_empresa, lAutomacaoBicoEmAcerto) Then
           CriaLogCupom "-- BOMBA LOCALIZADA OK -- CODIGO FISICO BOMBA= " & Format(Bomba.CodigoFisicoBomba, "00") & " BICO=" & Format(Bomba.Codigo, "00") & "TANQUE=" & Format(Bomba.NumeroTanque, "00")
           MovDocEletronicoItem.NumeroBomba = Format(Bomba.CodigoFisicoBomba, "00")
           MovDocEletronicoItem.NumeroBico = Format(Bomba.Codigo, "00")
           MovDocEletronicoItem.NumeroTanque = Format(Bomba.NumeroTanque, "00")
        Else
           CriaLogCupom "-- BOMBA NÃO LOCALIZADA --"
        End If

    End If
    
    MovDocEletronicoItem.Cancelado = False
    MovDocEletronicoItem.DataEntradaSaida = lData
    MovDocEletronicoItem.EtapaConcluida = Val(ETAPA_CONCLUIDA.GRAVADO)
    MovDocEletronicoItem.ProgramaOrigem = PROGRAMA_ORIGEM
    
    MovDocEletronicoItem.Periodo = lPeriodo
    
    'ME PARECE DESNECESSÁRIO AGORA - ALEX
    If MovDocEletronicoItem.CSTICMS = 60 Then
        MovDocEletronicoItem.ValorBCICMS = 0
        MovDocEletronicoItem.ValorBCICMSST = 0
    End If
'    If MovDocEletronicoItem.CSTIPI = 0 Then
'        MovDocEletronicoItem.ValorBcIpi = 0
'    End If
'    If MovDocEletronicoItem.CSTPIS = 4 Then
'        MovDocEletronicoItem.ValorBcPis = 0
'    End If
'    If MovDocEletronicoItem.CstCofins = 4 Then
'        MovDocEletronicoItem.ValorCofins = 0
'    End If

    
    GravaDocumentoEletronicoItem = MovDocEletronicoItem.Incluir
       
    Exit Function

FileError:
    Dim xString As String
    xString = "Numero=" & lNumeroCupom
    xString = xString & " - Ordem=" & lOrdem
    xString = xString & " - Data=" & lData
    xString = xString & " - Produto=" & CLng(txt_produto.Text)
    xString = xString & " - Quantidade=" & txt_quantidade.Text
    xString = xString & " - ValorUnitario=" & txt_valor_unitario.Text
    xString = xString & " - ValorTotal=" & txt_valor_total.Text
    Call CriaLogCupom("ERRO: Ao gravar Item do documento eletrônico - " & xString)
    Exit Function
End Function
Private Sub GravaLogEncerrantes()
    Dim i As Integer
    
    For i = 1 To lQtdBomba
        If EncerranteAtual.LocalizarCodigo(g_empresa, i) Then
            Call CriaLogAutomacaoEncerrante(lMarcaAutomacao, Date & "-" & Time & " - Bico:" & i & " - Encerrante:" & EncerranteAtual.Encerrante)
        Else
            Call CriaLogAutomacaoEncerrante(lMarcaAutomacao, Date & "-" & Time & " - Bico:" & i & " - *** NAO CADASTRADA ***")
        End If
    Next
    
End Sub
Function ValidaCliente() As Boolean
    ValidaCliente = False
    If dtcboCliente.BoundText = "" And (txt_cliente.Text <> "0" And txt_cliente.Text <> "00") Then
        MsgBox "Escolha o cliente.", vbInformation, "Atenção!"
        dtcboCliente.SetFocus
    Else
        ValidaCliente = True
    End If
End Function
Function ValidaCampos() As Boolean
    ValidaCampos = False
    If Not IsDate(lData) Then
        MsgBox "Não foi possível definir a data.", vbInformation, "Atenção!"
        txt_produto.SetFocus
    ElseIf lPeriodo = 0 Then
        MsgBox "Não foi possível definir o período.", vbInformation, "Atenção!"
        txt_produto.SetFocus
    ElseIf cboTipoSubEstoque.ListIndex = -1 Then
        MsgBox "Escolha o tipo de Sub-Estoque.", vbInformation, "Atenção!"
        cboTipoSubEstoque.SetFocus
    ElseIf dtcboCliente.BoundText = "" And (txt_cliente.Text <> "0" And txt_cliente.Text <> "00") Then
        MsgBox "Escolha o cliente.", vbInformation, "Atenção!"
        dtcboCliente.SetFocus
'    ElseIf Not Val(lNumeroCupom) > 0 Then
'        MsgBox "Não foi possível definir o número do cupom.", vbInformation, "Atenção!"
'        txt_produto.SetFocus
    ElseIf dtcboProduto.BoundText = "" Then
        MsgBox "Escolha o produto.", vbInformation, "Atenção!"
        dtcboProduto.SetFocus
    ElseIf Not fValidaValor4(txt_valor_unitario.Text) > 0 Then
        MsgBox "Informe o valor unitário do produto.", vbInformation, "Atenção!"
        txt_valor_unitario.SetFocus
    ElseIf Not fValidaValor(txt_quantidade.Text) > 0 Then
        MsgBox "Informe a quantidade.", vbInformation, "Atenção!"
        txt_quantidade.SetFocus
    ElseIf Not fValidaValor2(txt_valor_total.Text) > 0 Then
        MsgBox "Informe o valor total.", vbInformation, "Atenção!"
        txt_valor_total.SetFocus
    ElseIf Produto.CodigoGrupo <> lGrupoCombustivel And fValidaValor(txt_quantidade.Text) > lQtdMaxProduto Then
        MsgBox "Quantidade de produtos acima de " & Format(lQtdMaxProduto, "###,##0") & " não será aceita.", vbInformation, "Atenção!"
        txt_quantidade.SetFocus
    'ElseIf Produto.CodigoGrupo <> lGrupoCombustivel And CasaDecimalZerada(txt_quantidade.Text) = False Then
    ElseIf CasaDecimalZerada(txt_quantidade.Text) = False And PermiteValorFracionado(txt_produto.Text) = False Then
        MsgBox "Produto não será aceito com Quantidade fracionada.", vbInformation, "Informação Inconsistente!"
        txt_quantidade.SetFocus
    ElseIf lOrigemAutomacao = False And Produto.CodigoGrupo = lGrupoCombustivel And fValidaValor(txt_quantidade.Text) > lQtdMaxCombustivel Then
        MsgBox "Quantidade de combustíveis acima de " & Format(lQtdMaxCombustivel, "###,##0") & " não será aceita.", vbInformation, "Atenção!"
        txt_quantidade.SetFocus
    ElseIf g_automacao And lOrigemAutomacao = False And Produto.CodigoGrupo = lGrupoCombustivel Then
        MsgBox "Combustível somente será aceito pela automação.", vbInformation, "Atenção!"
        txt_produto.Text = ""
        dtcboProduto.BoundText = ""
        txt_valor_unitario.Text = ""
        txt_quantidade.Text = ""
        txt_valor_total.Text = ""
        txt_produto.SetFocus
    ElseIf Not ValidaCreditoCliente Then
        txt_quantidade.SetFocus
    ElseIf Not ValidaEstoque Then
        txt_quantidade.SetFocus
    Else
        ValidaCampos = True
    End If
End Function

Function PermiteValorFracionado(ByVal pCodigoProduto As String) As Boolean
     
    PermiteValorFracionado = False
     
    If Produto.CodigoGrupo = lGrupoCombustivel Then
        PermiteValorFracionado = True
        Exit Function
    End If
       
    If UCase(g_nome_empresa) Like "*MARQUES DE CASTRO & GABRIEL LTDA*" Then 'Esmeralda
         'If pCodigoProduto = "2824" Or pCodigoProduto = "2829" Then 'produtos a granel
         If pCodigoProduto = "779" Or pCodigoProduto = "3069" Or pCodigoProduto = "3081" Then 'produtos a granel
             PermiteValorFracionado = True
         End If
    ElseIf g_nome_empresa Like "*AUTO POSTO CLASSE A*" Then
         If pCodigoProduto = "2824" Or pCodigoProduto = "2829" Or pCodigoProduto = "3364" Then 'produtos a granel
             PermiteValorFracionado = True
         End If
    ElseIf UCase(g_nome_empresa) Like "*AUTO POSTO T13*" Then
         If pCodigoProduto = "10453" Or pCodigoProduto = "10192" Or pCodigoProduto = "703" Or pCodigoProduto = "10743" Then 'produtos a granel
             PermiteValorFracionado = True
         End If
    ElseIf UCase(g_nome_empresa) Like "*AUTO POSTO VERA CRUZ*" Then 'VERA CRUZ 2
         If pCodigoProduto = "1366" Or pCodigoProduto = "1435" Then  'produtos a granel
             PermiteValorFracionado = True
         End If
    End If
        
End Function

Private Function ProcessamentoComplementarEmAndamento() As Boolean
    Dim xProcessaComplementar As New cProcessaNFCeComplementar
    ProcessamentoComplementarEmAndamento = False
    
    If xProcessaComplementar.ExisteProcessamentoEmAndamento(g_empresa) Then
        ProcessamentoComplementarEmAndamento = True
        MsgBox "Processamento de NFCe Complementar em andamento. Aguarde", vbInformation, "Atenção!"
    End If
    
End Function

Function ValidaCampos2() As Boolean
    ValidaCampos2 = False
    If cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) >= 2 And cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) <= 3 Then
        If txt_numero_cheque.Text = "" Then
            MsgBox "Informe o número do cheque.", vbInformation, "Atenção!"
            txt_numero_cheque.SetFocus
        ElseIf txt_telefone.Text = "" Then
            MsgBox "Informe o número do telefone.", vbInformation, "Atenção!"
            txt_telefone.SetFocus
        Else
            ValidaCampos2 = True
        End If
    'ElseIf cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 5 Then
    '    If l_codigo_cliente = 0 Then
    '        MsgBox "Forma de pagamento exclusiva para cliente.", vbInformation, "Atenção!"
    '        cbo_forma_pagamento.SetFocus
    '    ElseIf txt_numero_nota_abastecimento.Text = "" Then
    '        MsgBox "Informe o número da nota de abastecimento.", vbInformation, "Atenção!"
    '        txt_numero_nota_abastecimento.SetFocus
    '    Else
    '        ValidaCampos2 = True
    '    End If
    'ElseIf cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 5 And l_codigo_cliente = 0 Then
    '    MsgBox "A forma de pagamento escolhida só é permitida para cliente cadastrado.", vbInformation, "Atenção!"
    '    cbo_forma_pagamento.SetFocus
    ElseIf fValidaValor(txt_valor_recebido.Text) < fValidaValor(lbl_valor_compra.Caption) Then
        MsgBox "O valor recebido não pode ser menor que " & lbl_valor_compra.Caption & ".", vbInformation, "Atenção!"
        txt_valor_recebido.SetFocus
    ElseIf Not ValidaCnpjCpf Then
        txt_cpf.SetFocus
    ElseIf fValidaValor(txt_valor_recebido.Text) > fValidaValor(lbl_valor_compra.Caption) And (cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 4 Or cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 6 Or cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 7 Or cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 8 Or cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 9 Or cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 10 Or cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 11 Or cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 12 Or cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 13 Or cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 14 Or cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 15 Or cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 17) Then
        MsgBox "O valor recebido não deve ser maior que o valor da compra!", vbInformation, "Valor não aceito!"
        txt_valor_recebido.SetFocus
    ElseIf lBloqueiaDesconto And fValidaValor(txt_valor_desconto.Text) > 0 And lValorDescontoConcedido = 0 Then
        MsgBox "Empresa não configurada para desconto!", vbInformation, "Valor não aceito!"
        txt_valor_desconto.Text = "0,00"
        If txt_valor_desconto.Enabled = True Then
            txt_valor_desconto.SetFocus
        Else
            txt_valor_recebido.SetFocus
        End If
    ElseIf Not ValidaFormaPagamentoPOS Then
        cbo_forma_pagamento.SetFocus
    ElseIf Not ValidaCartaoPOS Then
        cbo_forma_pagamento.SetFocus
    ElseIf ValidaCreditoCliente = False Then
    Else
        ValidaCampos2 = True
    End If
End Function
Function ValidaFormaPagamentoPOS() As Boolean
'    ValidaFormaPagamentoPOS = False
'    If cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 16 Then
'        If ConfiguracaoDiversa.LocalizarCodigo(1, "Empresa Azure para LIO") Then
'            If ConfiguracaoDiversa.Codigo = 96 Then
'                MsgBox "Selecione outra forma de pagamento.", vbInformation, "Forma de Pagamento Não Aceita!"
'                cbo_forma_pagamento.SetFocus
'                ValidaFormaPagamentoPOS = False
'            Else
'                ValidaFormaPagamentoPOS = True
'            End If
'        Else
'            ValidaFormaPagamentoPOS = True
'        End If
'    Else
        ValidaFormaPagamentoPOS = True
'    End If
End Function
Function ValidaCamposPonto() As Boolean
    ValidaCamposPonto = False
    If Val(dtcboFuncionario.BoundText) = 0 Then
        MsgBox "Selecione o funcionário.", vbInformation, "Atenção!"
        dtcboFuncionario.SetFocus
    Else
        ValidaCamposPonto = True
    End If
End Function
Function ValidaCreditoCliente() As Boolean
    Dim xValorCupomAtual As Currency
    ValidaCreditoCliente = True
    
    
    
    xValorCupomAtual = 0
    If lNumeroCupom = lNumeroUltimoCupom And lOrdem > 1 Then
        xValorCupomAtual = lTotalCupom
    End If
    xValorCupomAtual = xValorCupomAtual + fValidaValor2(txt_valor_total.Text)
    
    If g_nivel_acesso >= lRestringeVendaCredito And dtcboCliente.Text <> "" Then
        If Cliente.GeraNotaAbastecimento Then
            If VerificaSaldoCreditoCliente(xValorCupomAtual) = False Then
                ValidaCreditoCliente = False
            End If
        End If
    End If
End Function
Function ValidaEstoque() As Boolean
    ValidaEstoque = False
    If lBloqueiaEstoque = False And lBloqueiaSubEstoque = False Then
        ValidaEstoque = True
        Exit Function
    End If
    If lBaixaAutomaticaNoEstoque = False Then
        ValidaEstoque = True
        Exit Function
    End If
    If Produto.CodigoGrupo = lGrupoCombustivel Then
        ValidaEstoque = True
        Exit Function
    End If
'
'  Aqui foi comentado porque no posto Pelicano não tem controle do que tem na pista
'
'    If Not SubEstoque.LocalizarCodigo(g_empresa, CLng(txt_produto.Text), lTipoMovimento) Then
'        MsgBox "SubEstoque não cadastrado.", vbInformation, "Erro de Verificação!"
'        txt_produto.SetFocus
'        Exit Function
'    Else
'        If SubEstoque.Quantidade < fValidaValor(txt_quantidade.Text) Then
'            MsgBox "Não é permitido tirar cupom fiscal acima da quantidade em estoque." & Chr(10) & "A quantidade atual no SubEstoque é: " & Format(SubEstoque.Quantidade, "##,###,##0.00") & ".", vbInformation, "Estoque Insuficiente!"
'        Else
'            ValidaEstoque = True
'        End If
'    End If
    If Not Estoque.LocalizarCodigo(g_empresa, CLng(txt_produto.Text)) Then
        MsgBox "Estoque não cadastrado.", vbInformation, "Erro de Verificação!"
        txt_produto.SetFocus
        Exit Function
    Else
        If Estoque.Quantidade < fValidaValor(txt_quantidade.Text) Then
            MsgBox "Não é permitido tirar cupom fiscal acima da quantidade em estoque." & Chr(10) & "A quantidade atual no Estoque é: " & Format(Estoque.Quantidade, "##,###,##0.00") & ".", vbInformation, "Estoque Insuficiente!"
        Else
            ValidaEstoque = True
        End If
    End If
End Function
Function ValidaCartaoPOS() As Boolean
    Dim xNomeBandeira As String
    Dim xCodigoCartao As Integer
    Dim xCodigoCartaoDesconto As Integer
    
    ValidaCartaoPOS = False
    lNFCe_tPag = ""
    lNFCe_vPag = 0
    lNFCe_TpIntegra = 0
    lNFCe_CNPJCartao = ""
    lNFCe_tBand = ""
    lNFCe_cAut = ""
    xCodigoCartaoDesconto = 0
    lPOSImprimeNFCe = True
    If cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 16 Then
        
        If lDescontoCartaoEspecial = True Then
            If ConfiguracaoDiversa.LocalizarCodigo(1, "CODIGO CARTAO ESPECIAL DESCONTO") Then
                If ConfiguracaoDiversa.Codigo > 0 Then
                    xCodigoCartaoDesconto = ConfiguracaoDiversa.Codigo
                End If
            End If
        End If
    
        g_string = "Movimento_NFCe_Auto|@|Incluir|@|" & Format(lDataCupom, "dd/MM/yyyy") & "|@|" & lPeriodo & "|@|" & lTipoMovimento & "|@|" & lIlha & "|@|1|@|2|@|" & g_usuario & "|@|" & lbl_valor_compra.Caption & "|@|" & lNumeroCupom & "|@|" & l_codigo_funcionario & "|@|" & xCodigoCartaoDesconto & "|@|"
        movimento_cartao_credito.Show 1
        If RetiraGString(1) = "Retorno-Movimento_NFCe_Auto" Then
        
            'xTipoRetorno = RetiraGString(1)            'Retorno-Movimento_NFCe_Auto
            xCodigoCartao = Val(RetiraGString(2))       'Codigo da Cartao
            xNomeBandeira = RetiraGString(3)            'Nome Bandeira Cartao
            
            If xCodigoCartaoDesconto > 0 And xCodigoCartaoDesconto <> xCodigoCartao Then
                g_string = ""
                MsgBox "Cartão selecionado não é valido para desconto!", vbInformation, "Dados Inconsistentes!"
                Exit Function
            End If

            
            lNFCe_vPag = fValidaValor(RetiraGString(4)) 'Valor
            lNFCe_cAut = RetiraGString(5)               'Numero da Autorização
            lNFCe_tPag = "03"                           '03-Cartão Cédito
            If UCase(xNomeBandeira) Like "*DÉBITO*" Or UCase(xNomeBandeira) Like "*DEBITO*" Then
                lNFCe_tPag = "04"                       '04-Cartão Débito
            End If
            lNFCe_TpIntegra = 2                         '1-TEF, 2-POS
            
            'Define a BANDEIRA
            lNFCe_tBand = DefineNFCe_tBand(xNomeBandeira)
            
            'Define a INTEGRADORA
            lNFCe_CNPJCartao = DefineNFCe_CNPJCartao(xNomeBandeira)
            
            If xCodigoCartaoDesconto > 0 And lDescontoCartaoEspecial = True Then
                lIntegraDescontoCartaoEspecial = True
            End If
            
            If ConfiguracaoDiversa.LocalizarCodigo(1, "NFCEAUTO: Cartao POS Imprime NFCe") Then
                lPOSImprimeNFCe = ConfiguracaoDiversa.Verdadeiro
            End If
            
            ValidaCartaoPOS = True
        Else
            MsgBox "Favor lançar o comprovante de venda do Cartão Primeiro!", vbInformation, "Dados Incompleto!"
        End If
        g_string = ""
    Else
        ValidaCartaoPOS = True
    End If
End Function
Private Sub DefineCartaoTefParaNFCe()
    Dim xNomeBandeira As String
    
    lNFCe_tPag = ""
    lNFCe_vPag = 0
    lNFCe_TpIntegra = 0
    lNFCe_CNPJCartao = ""
    lNFCe_tBand = ""
    lNFCe_cAut = ""
    
    CriaLogCupom ("")
    
    xNomeBandeira = CartaoCredito.Nome                 'Nome Bandeira Cartao
    lNFCe_vPag = fValidaValor(txt_valor_recebido.Text) 'Valor
    
    If (lCartaoAutorizacao = Empty) Then
        lNFCe_cAut = lNSU ' Como a variável da autorização (POSIÇÃO 013) está em branco utilizo o NSU (POSIÇÃO 012)
    Else
        lNFCe_cAut = lCartaoAutorizacao 'Numero da Autorização
    End If
    
    
    lNFCe_tPag = "03"                                  '03-Cartão Cédito
    If UCase(xNomeBandeira) Like "*DÉBITO*" Or UCase(xNomeBandeira) Like "*DEBITO*" Then
        lNFCe_tPag = "04"                       '04-Cartão Débito
    End If
    lNFCe_TpIntegra = 2                         '1-TEF, 2-POS
    
    'Define a BANDEIRA
    lNFCe_tBand = DefineNFCe_tBand(xNomeBandeira)
    
    'Define a INTEGRADORA
    lNFCe_CNPJCartao = DefineNFCe_CNPJCartao(xNomeBandeira)
    
    Call CriaLogCupom("DefineCartaoTefParaNFCe - DEU CERTO OK - xNomeBandeira=" & xNomeBandeira & " - lNFCe_vPag=" & lNFCe_vPag & " - lNFCe_TpIntegra=" & lNFCe_TpIntegra & " - lNFCe_CNPJCartao=" & lNFCe_CNPJCartao & " - lNFCe_tBand=" & lNFCe_tBand)

End Sub
Private Function DefineNFCe_tBand(ByVal pNomeBandeira As String) As String
    DefineNFCe_tBand = "99"
    If UCase(pNomeBandeira) Like "*VISA*" Or UCase(pNomeBandeira) Like "*ELETRON*" Then
        DefineNFCe_tBand = "01"                      'Bandeira 01-Visa
    ElseIf UCase(pNomeBandeira) Like "*MASTER*" Or UCase(pNomeBandeira) Like "*MAESTRO*" Then
        DefineNFCe_tBand = "02"                      'Bandeira 02-Mastercard
    ElseIf UCase(pNomeBandeira) Like "*AMERICAN*" Then
        DefineNFCe_tBand = "03"                      'Bandeira 03-American Express
    ElseIf UCase(pNomeBandeira) Like "*SOROCRED*" Or UCase(pNomeBandeira) Like "*SORO CRED*" Then
        DefineNFCe_tBand = "04"                      'Bandeira 04-Sorocred
    End If
End Function
Private Function DefineNFCe_CNPJCartao(ByVal pNomeBandeira As String) As String
    DefineNFCe_CNPJCartao = "01027058000191"
    If UCase(pNomeBandeira) Like "*CIELO*" Then
        DefineNFCe_CNPJCartao = "01027058000191"     'CIELO
    ElseIf UCase(pNomeBandeira) Like "*REDECARD*" Or UCase(pNomeBandeira) Like "*REDE*" Then
        DefineNFCe_CNPJCartao = "01425787000101"     'REDECARD
    ElseIf UCase(pNomeBandeira) Like "*AMERICAN*" Or UCase(pNomeBandeira) Like "*AMEX*" Then
        DefineNFCe_CNPJCartao = "60419645000195"     'AMERICAN EXPRESS
    ElseIf UCase(pNomeBandeira) Like "*SOROCRED*" Or UCase(pNomeBandeira) Like "*SORO CRED*" Then
        DefineNFCe_CNPJCartao = "60114865000100"     'SOROCRED
    ElseIf UCase(pNomeBandeira) Like "*BANCOOB*" Then
        DefineNFCe_CNPJCartao = "02038232000164"     'BANCOOB
    End If



End Function
Private Function ValidaCnpjCpf() As Boolean
    Dim xCnpjCpf As String
    
    xCnpjCpf = fDesmascaraNumeroString(txt_cpf.Text)
    ValidaCnpjCpf = False
    If Len(xCnpjCpf) = 0 Then
        ValidaCnpjCpf = True
    ElseIf Len(xCnpjCpf) = 11 Then
        If CalculaDigitoCPF(xCnpjCpf) Then
            ValidaCnpjCpf = True
        Else
            MsgBox "CPF Inválido, conforme cálculo de dígito.", vbCritical, "Dados Inválido!"
        End If
    ElseIf Len(xCnpjCpf) = 14 Then
        If CalculaDigitoCNPJ(xCnpjCpf) Then
            ValidaCnpjCpf = True
        Else
            MsgBox "CNPJ Inválido, conforme cálculo de dígito.", vbCritical, "Dados Inválido!"
        End If
    Else
        ValidaCnpjCpf = False
        MsgBox "Quantidade de dígitos inválido para CNPJ ou CPF.", vbCritical, "Dados Inválido!"
    End If
End Function
Private Sub cmdPesquisa_Click()
    Dim xString As String
    Dim xCodigo As Long
    
    xString = g_string
    'True para ocultar algumas colunas da pesquisa
    g_string = "True|@|"
    consulta_produto.Show 1
    If Len(g_string) > 0 Then
        xCodigo = RetiraGString(1)
        If Produto.LocalizarCodigo(xCodigo) Then
            txt_produto.Text = xCodigo
            dtcboProduto.BoundText = CLng(txt_produto.Text)
            txt_produto_LostFocus
        End If
    End If
    g_string = xString
End Sub
Private Sub dtcboCliente_GotFocus()
    lOrigemFocus = "dtcboCliente"
    l_mensagem = Space(165) & "Selecione o cliente."
End Sub
Private Sub dtcboCliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_produto.SetFocus
    End If
End Sub
Private Sub dtcboCliente_LostFocus()
    If dtcboCliente.BoundText <> "" Then
        l_codigo_cliente = Val(dtcboCliente.BoundText)
        If Cliente.LocalizarCodigo(CLng(dtcboCliente.BoundText)) Then
            txt_cliente.Text = Cliente.Codigo
            If txt_produto.Enabled Then
                txt_produto.SetFocus
            End If
            Exit Sub
        End If
    ElseIf txt_cliente.Text = "0" Or txt_cliente.Text = "00" Then
        If frmDados.Enabled = True Then
            txt_produto.SetFocus
        End If
    End If
End Sub
Private Sub dtcboFuncionario_GotFocus()
    lOrigemFocus = "dtcboFuncionario"
    l_mensagem = Space(165) & "Selecione o funcionário."
End Sub
Private Sub dtcboFuncionario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_senha_ponto.SetFocus
    End If
End Sub
Private Sub dtcboFuncionario_LostFocus()
    If dtcboFuncionario.BoundText <> "" Then
        txt_funcionario_ponto = dtcboFuncionario.BoundText
        txt_funcionario_ponto_LostFocus
        If txt_senha_ponto.Visible Then
            txt_senha_ponto.SetFocus
        End If
    End If
End Sub
Private Sub dtcboProduto_GotFocus()
    lOrigemFocus = "dtcboProduto"
    l_mensagem = Space(165) & "Selecione o produto."
End Sub
Private Sub dtcboProduto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_quantidade.SetFocus
    End If
End Sub
Private Sub dtcboProduto_LostFocus()
    If dtcboProduto.BoundText <> "" Then
        txt_produto.Text = dtcboProduto.BoundText
        
        If ConfiguracaoDiversa.LocalizarCodigo(1, "ECF: BLOQUEIA PRODUTO-" & Format(CLng(txt_produto.Text), "0000")) Then
            If ConfiguracaoDiversa.Verdadeiro = True Then
                MsgBox "Produto Bloqueado para emissão de cupom fiscal.", vbInformation, "Produto Bloqueado!"
                txt_valor_unitario.Text = ""
                txt_produto.SetFocus
                Exit Sub
            End If
        End If
        
        
        If Produto.LocalizarCodigo(CLng(txt_produto.Text)) Then
            If lExigeNCM = True Then
                If LocalizarNCM(0, Trim(Produto.CodigoNCM)) = False Then
                    txt_produto.SetFocus
                    Exit Sub
                End If
            End If
            If Estoque.LocalizarCodigo(g_empresa, CLng(txt_produto.Text)) Then
                If Not Aliquota.LocalizarCodigo(lSerieECF, Produto.CodigoAliquota) Then
                    MsgBox "Aliquota não cadastrada!", vbInformation, "Erro de Integridade!"
                    txt_produto.SetFocus
                    Exit Sub
                End If
                If Estoque.PrecoVenda <> 0 Then
                    txt_valor_unitario.Text = Format(Estoque.PrecoVenda, "###,##0.0000")
                Else
                    txt_valor_unitario.Text = Format(Produto.PrecoVenda, "###,##0.0000")
                End If
            Else
                MsgBox "Estoque não cadastrado.", vbInformation, "Erro de Verificação!"
                txt_valor_unitario.Text = ""
                txt_valor_unitario.SetFocus
                Exit Sub
            End If
            
            'LUIZ 28/11/2016 não é obrigatorio informar o cest
'            If (Aliquota.CodigoFiscal = CODIGO_FISCAL_SUBSTITUICAO) And Trim(Produto.CEST) = Empty Then
'                MsgBox "O produto " & Trim(Produto.Nome) & vbCrLf & "é Substituição Tributária e está cadastrado sem CEST.", vbInformation, "Produto Sem CEST!"
'                txt_produto.SetFocus
'                Exit Sub
'            End If
            
            
        End If
        txt_quantidade.SetFocus
    Else
        'If txt_produto = "" And l_flag_cupom_fiscal = "A" Then
        '    CancelaCupom
        'End If
    End If
End Sub
Private Sub Form_Activate()
    'MsgBox "Activate iniciado"
    'Call CriaLogECF(Date & " " & Time & " Activate: N.Cupom=" & lNumeroCupom & " - Retorno lRespostaTEF=" & lRespostaTEF)
    Call CriaLogECF(Date & " " & Time & " Form_Activate Iniciado")
    If Not TestaEmpresa Then
        lFinalizaAutomatico = True
        Unload Me
        Screen.MousePointer = 1
        Exit Sub
    End If
    If fUsaNFCe = False Then
        MsgBox "Este computador não está configurado para emissão de NFCe!", vbInformation, "Operação não autorizada!"
        lFinalizaAutomatico = True
        Unload Me
        Screen.MousePointer = 1
        Exit Sub
    End If
    
    If g_empresa <> lEmpresa Then
        flag_Movimento_Cupom_Fiscal = 0
    End If
    Call CriaLogECF(Date & " " & Time & " Activate: *** TESTE DE PERFORMANCE 10 ***")
    
    'Sempre que voltar abre a porta, pelo motivo que quando entra no caixa de pista
    'pra imprimir vale abastecimento, a porta tem que ser fechada pra dar certo
    'If lImpBematech Then
        'Call CriaLogCupom("Bematech_FI_AbrePortaSerial")
        'BemaRetorno = Bematech_FI_AbrePortaSerial()
        'Call CriaLogCupom("Bematech_FI_AbrePortaSerial - BemaRetorno=" & BemaRetorno)
    'End If
    Call CriaLogECF(Date & " " & Time & " Activate: *** TESTE DE PERFORMANCE 20 ***")
    
    If flag_Movimento_Cupom_Fiscal = 0 Then
        Call CriaLogCupom(Date & " " & Time & " Activate: flag_Movimento_Cupom_Fiscal = 0")
        'Call TestaConexao(1, "Form.Activate")
        'Set CerradoTef = Nothing
        'Call TestaConexao(2, "Form.Activate")
        Set CerradoTef = New CerradoComponenteTef
        'Call TestaConexao(3, "Form.Activate")
        'MsgBox "CerradoTef.VerificaPendencia iniciado"
        
        Call CerradoTef.VerificaPendencia("NFCe")
        'Call TestaConexao(4, "Form.Activate")
        Set CerradoTef = Nothing
        'Call TestaConexao(5, "Form.Activate")
        
        'MsgBox "CerradoTef.VerificaPendencia iniciado"
        
        
        
        
'        MsgBox "TESTE ALEX BRUTO 1"
'        lSerieECF = ReadINI("CUPOM FISCAL", "Serie ECF", gArquivoIni)
'        lNumeroPDV = 1
'        If ECF.LocalizarNumeroSerie(g_empresa, lSerieECF) Then
'            MsgBox "TESTE ALEX BRUTO 1B"
'            lNumeroPDV = ECF.Codigo
'        End If
'        MsgBox "TESTE ALEX BRUTO 2"
        
        
        
        

        Call CriaLogECF(Date & " " & Time & " Activate: *** TESTE DE PERFORMANCE 20-a ***")
        AtualizaConstantes
        Call CriaLogECF(Date & " " & Time & " Activate: *** TESTE DE PERFORMANCE 20-b ***")
        
        
        
        
        
        
'        MsgBox "TESTE ALEX BRUTO 3"
'        lSerieECF = ReadINI("CUPOM FISCAL", "Serie ECF", gArquivoIni)
'        lNumeroPDV = 1
'        If ECF.LocalizarNumeroSerie(g_empresa, lSerieECF) Then
'            MsgBox "TESTE ALEX BRUTO 3B"
'            lNumeroPDV = ECF.Codigo
'        End If
'        MsgBox "TESTE ALEX BRUTO 4"
        
        
        
        
        
        
        
        
        
        
        PosicionamentoInicialBombas
        Call CriaLogECF(Date & " " & Time & " Activate: *** TESTE DE PERFORMANCE 20-c ***")
        
        lEmpresa = g_empresa
        
        BuscaDados
        Call CriaLogECF(Date & " " & Time & " Activate: *** TESTE DE PERFORMANCE 20-d ***")
        
        
        Screen.MousePointer = 1
        frm_ponto.ZOrder 0
        txt_funcionario_ponto = ""
        dtcboFuncionario.BoundText = 0
        txt_senha_ponto.Text = ""
        frmDados.Enabled = False
        frm_fechamento_cupom.Enabled = False
        txt_cupom_fiscal.Enabled = False
        If txt_funcionario_ponto.Enabled Then
            txt_funcionario_ponto.SetFocus
        End If
        If lRestringeDescontoNivelUsuario Then
            txt_valor_desconto.Enabled = False
        End If
'        If lImpBematech Then
'            If ReadINI("CUPOM FISCAL", "Grava CAT52", gArquivoIni) = "NAO" Then
'            Else
'                Call LoopGravaCat52(CDate(Date - 1), CDate(Date - 1))
'            End If
'        End If
    Else
        flag_Movimento_Cupom_Fiscal = 0
    End If
    Call CriaLogECF(Date & " " & Time & " Activate: *** TESTE DE PERFORMANCE 30 ***")
    'NAO CHAMAR MAIS BAIXA DE ABASTECIMENTO, O PROGRAMA DE AUTOMACAO JA FAZ ISSO
    'BaixaAbastecimentoAcertado

   ' AtivaReducaoZ
    
    Dim xFlag As Integer
    Call CriaLogECF(Date & " " & Time & " Activate: *** TESTE DE PERFORMANCE 40 ***")
    If lImpBematech Then
        'Call CriaLogCupom("Bematech_FI_FlagsFiscais(xFlag)")
        'Bematech_FI_FlagsFiscais(xFlag) 'NFCE - ALEX
        'Call CriaLogCupom("Bematech_FI_FlagsFiscais(xFlag) - xFlag=" & xFlag & " - BemaRetorno=" & BemaRetorno)
        'Call CriaLogECF(Date & " " & Time & " Bematech_FI_FlagsFiscais: xFlag=" & xFlag)
    End If
    Call CriaLogECF(Date & " " & Time & " Activate: *** TESTE DE PERFORMANCE 50 ***")
'    If xFlag = ::: Then
'        If MsgBox("Redução Z do dia anterior pendente!" & Chr(10) & "Deseja imprimir a redução Z pendente?", vbYesNo + vbQuestion + vbDefaultButton2, "Redução Z Pendente!") = vbYes Then
'            Call ImprimeReducaoZ
'            Call CriaLogECF(Date & " " & Time & " Redução Z impressa automaticamente pelo SGP")
'        End If
'    End If
'MsgBox "Activate Finalizado"


    'Codigo abaixo pra debugar cartao nao integrado.
    'Precisa do arquivo de retorno estar na pasta c:\vb5\sgp\data\teste.txt
'    MsgBox "VERSAO PRA DEBUG"
'    If IntegraCartaoCreditoNoCaixa Then
'        DefineCartaoTefParaNFCe
'        AtualizaTabelaCartaoCredito
'        ' Desconto Cartao Correios
'        If lIntegraDescontoCartaoCorreios = True And lValorDescontoConcedido > 0 Then
'            If Not IncluiMovimentoCaixa(lDataCupom, lPeriodo, True, "DescontoCartaoCorreios", lValorDescontoConcedido, "", "Cartão Desconto Correios") Then
'                Call CriaLogCupom("Erro cmd_DescontoCorreio_Click:Desconto Cartao Correios não integrada no caixa.")
'                MsgBox "Não foi possível integrar Desconto Cartão Correios no caixa!", vbInformation, "Erro de Integridade!"
'            End If
'        End If
'    Else
'        DefineCartaoTefParaNFCe
'    End If
End Sub
Private Sub Form_Deactivate()
    flag_Movimento_Cupom_Fiscal = 1
End Sub
Private Sub Form_Load()
    
    Dim xEmpresa As New cEmpresa
    lUfEmpresa = ""
 
' MsgBox "Form_Load iniciado- TESTE ALEX"
'    If ConfiguracaoDiversa.LocalizarCodigo(1, "ECF: PULAR X LINHAS ENTRE CV") Then
'        lLinhasEntreCV = ConfiguracaoDiversa.Codigo
'    End If
'
' MsgBox "Form_Load iniciado- TESTE ALEX"
    Dim xTipoVenda As String
    Dim xString As String
    
    Call GravaAuditoria(1, Me.name, 1, "")
    
    'Rotina criada para excluir baixa de abastecimento duplicada
    'zzExcluiBaixaDuplicada
    Call PreencheDicionarioCSTPisCofins
    
        
    If xEmpresa.LocalizarCodigo(g_empresa) Then
        lUfEmpresa = xEmpresa.Estado 'utizado no timer que atualiza bombas automaticamente, será ativado somente para estado do Pará
    End If
        
    Call DefinePortaEcf
    Call ConsisteNFCePendentes(5)
    lLoja = False
    x_tempo = 0
    Me.Left = (Screen.Width - Me.Width)
    Me.Top = 0
    lFinalizaAutomatico = False
    lCodigoFiscal = "  "
    frm_fechamento_cupom.Left = 120
    frmDescarregar.Enabled = False
    frmDescarregar.Visible = False
    btnAbortaNFCe.Enabled = False
    
    TimerAguarde.Interval = 0
    TimerAguarde.Enabled = False
    frameAguarde.ZOrder 1
    
    lOrigemFocus = ""
    lEcfTruncamento = False
    lEcfQtdCasasDecimais = 3
    'lTestaReducaoZpendente = True
    lLocalImpressao = 1 '0-Video, 1-Impressora
    
    lTEF = False
    lRespostaTEF = False
    lNumeroCupom = 0
    'Call TestaConexao(1, "Form.Load")
    Set CerradoTef = New CerradoComponenteTef
    
    'MsgBox "CerradoTef.VerificaPendencia iniciado- TESTE ALEX"
    'Call TestaConexao(2, "Form.Load")
    Call CerradoTef.VerificaPendencia("NFCe")
    'Call TestaConexao(3, "Form.Load")
    Set CerradoTef = Nothing
    'Call TestaConexao(4, "Form.Load")
   ' MsgBox "CerradoTef.VerificaPendencia FINALIZADO- TESTE ALEX"
    
    'AtivaReducaoZ
    
    MovimentoBomba.NomeTabela = "Movimento_Bomba_Cupom"
    MovimentoBombaEscritorio.NomeTabela = "Movimento_Bomba"
    PreencheCboFormaPagamento
    PreencheCboBico
    PreencheCboTipoSubEstoque
    
    'MsgBox "EXECUTA Conectar.RsConexao(SELECT Codigo, [Razao Social] FROM Cliente WHERE Inativo =  & preparaBooleano(False) &  ORDER BY [Razao Social])- TESTE ALEX"
    
    Set adodcCliente.Recordset = Conectar.RsConexao("SELECT Codigo, [Razao Social] FROM Cliente WHERE Inativo = " & preparaBooleano(False) & " ORDER BY [Razao Social]")
    
    xTipoVenda = ReadINI("CUPOM FISCAL", "Tipo de Venda", gArquivoIni)
    
    'MsgBox "TIPO DE VENDA: " & xTipoVenda & "TESTE ALEX"
    
    lSQL = "SELECT Codigo, Nome"
    lSQL = lSQL & "  FROM Produto"
    lSQL = lSQL & " WHERE Inativo = " & preparaBooleano(False)
    lSQL = lSQL & " AND [Imprime Cupom Fiscal] = " & preparaBooleano(True)
    If xTipoVenda = "AUTOMACAO/CONVENIENCIA" Then
        lLoja = True
        'lSQL = lSQL & "   AND [Exclusivo Loja] = " & preparaBooleano(True)
    Else
        lSQL = lSQL & "   AND [Exclusivo Posto] = " & preparaBooleano(True)
    End If
    
    'MsgBox "Configuracao.LocalizarCodigo(g_empresa) - TESTE ALEX"
      
    If Configuracao.LocalizarCodigo(g_empresa) Then
        If Mid(Configuracao.OutrasConfiguracoes, 8, 1) = "N" Then
            lSQL = lSQL & "   AND Unidade <> " & preparaTexto("SRV")
        End If
    End If
    lSQL = lSQL & " ORDER BY Nome"
    Set adodcProduto.Recordset = Conectar.RsConexao(lSQL)
    
    'MsgBox "TESTE ALEX 4"
    
    Set adodcFuncionario.Recordset = Conectar.RsConexao("Select Codigo, Nome From Funcionario WHERE Empresa = " & g_empresa & " AND Situacao = " & preparaTexto("A") & " ORDER BY [Nome]")
    l_flag_cupom_fiscal = "F"
    lNotificacaoGic = False
    If ConfiguracaoDiversa.LocalizarCodigo(1, "GIC: Notificacao Periodica") Then
        If ConfiguracaoDiversa.Verdadeiro Then
            lNotificacaoGic = True
            menu_personalizado.AtivaVerificacaoGIC
        End If
    End If
    
    'MsgBox "TESTE ALEX 5"
    
    cmdIniciaProcessoFinalizacaoNFCe.Enabled = False
    Call AbilitaMenu(False)
    TestaCupomDemonstracao
    'Primeiro comando no ECF
    
    'MsgBox "TESTE ALEX 6"
    
    'ImprimeProgramaFormaPagamento
    lNumeroUltimoCupom = 0
    lTotalCupom = 0
    lDescontoEspecial = 0
    
    'NFCE - ALEX - 03-08-2016
'    If lExisteImpressora = False And lCupomDemonstracao = False Then
'        MsgBox "Problemas de comunicação com a Impresão." & Chr(13) & "Não será possível imprimir cupom fiscal.", vbCritical, "Erro de Comunicação!"
'        Finaliza
'        End
'    End If
    lAutomacaoFlagVendaAutomatica = False
    
    'MsgBox "TESTE ALEX 7"
    
    lNomeArquivoAutomacaoIni = ArquivoAutomacaoIni
    If ReadINI("OUTRAS", "Apenas Visualizar", lNomeArquivoAutomacaoIni) = "NAO" Then
        If lMarcaAutomacao = "COMPANY" Then
            If ComunicaAutomacaoCerradoBD("AUTOMACAO ATIVADA", "") = False Then
                If (MsgBox("Não foi possível comunicar com o progama AutoCerradoCompany." & vbCrLf & vbCrLf & "Se este programa não for aberto, " & vbCrLf & "algumas funcionalidade não serão executadas." & vbCrLf & vbCrLf & "Deseja abrir programa de comunicação com o equipamento de Automação?", vbQuestion + vbYesNo + vbDefaultButton1, "Deseja Abrir Programa Cerrado Automação?") = vbYes) Then
                    LoopAbreAutomacao
                Else
                End If
            End If
        ElseIf lMarcaAutomacao = "HOROUSTECH" Then
            If ComunicaAutomacaoCerradoBD("AUTOMACAO ATIVADA", "") = False Then
                If (MsgBox("Não foi possível comunicar com o progama AutoCerradoHorousTech." & vbCrLf & vbCrLf & "Se este programa não for aberto, " & vbCrLf & "algumas funcionalidade não serão executadas." & vbCrLf & vbCrLf & "Deseja abrir programa de comunicação com o equipamento de Automação?", vbQuestion + vbYesNo + vbDefaultButton1, "Deseja Abrir Programa Cerrado Automação?") = vbYes) Then
                    LoopAbreAutomacao
                Else
                End If
            End If
        ElseIf lMarcaAutomacao = "EZTECH" Then
            If ComunicaAutomacaoCerradoBD("AUTOMACAO ATIVADA", "") = False Then
                If (MsgBox("Não foi possível comunicar com o progama AutoCerradoEZ." & vbCrLf & vbCrLf & "Se este programa não for aberto, " & vbCrLf & "algumas funcionalidade não serão executadas." & vbCrLf & vbCrLf & "Deseja abrir programa de comunicação com o equipamento de Automação?", vbQuestion + vbYesNo + vbDefaultButton1, "Deseja Abrir Programa Cerrado Automação?") = vbYes) Then
                    LoopAbreAutomacao
                Else
                End If
            End If
        End If
    End If
    
    'MsgBox "TESTE ALEX 8"
    
    If lImpBematech Then
        TestaEncerramentoCupomFiscal
    ElseIf lImpQuick Then
        If EcfQuickLeRegistrador("EstadoFiscal", "Long", 5) = "2" Then
            Dim xValor As Currency
            'Cancela o cupom aberto
            Call EcfQuickCancelaCupom
            'nao pode fechar como dinheiro
            'xValor = fValidaValor(EcfQuickLeRegistrador("TotalDocLiquido", "Monetario", 6))
            'If Not EcfQuickPagaCupom(0, "Dinheiro", "** Fechado Automaticamente pelo Sistema **", xValor) Then
            '    MsgBox "Erro ao pagar cupom fiscal na Ecf Quick", vbCritical, "Erro ao Finalizar Cupom"
            'End If
        End If
        
        'teste para fechar gerencial caso esteja aberto
        If EcfQuickLeRegistrador("EstadoFiscal", "Long", 5) = "32" Then
            Call EcfQuickEncerraDocumento(0, "Gerencial")
        End If
        If EcfQuickLeRegistrador("EstadoFiscal", "Long", 5) = "64" Then
            Call EcfQuickEncerraDocumento(0, "Gerencial")
        End If
        
        If EcfQuickLeRegistrador("EstadoFiscal", "Long", 5) = "8" Then
            If Not EcfQuickEncerraDocumento(l_nome_funcionario, "Cerrado Informatica (62) 3277-1017") Then
                MsgBox "Erro ao finalizar cupom fiscal na Ecf Quick", vbCritical, "Erro ao Finalizar Cupom"
            End If
        End If
    ElseIf lImpDaruma Then
        'xString = Space(2)
        'BemaRetorno = Daruma_FI_StatusCupomFiscal(xString)
        'If Mid(xString, 1, 1) = "1" Then
        '    DarumaBuscaRetorno
        '    xString = Space(18)
        '    BemaRetorno = Daruma_FI_SaldoAPagar(xString)
        '    BemaRetorno = Daruma_FI_IniciaFechamentoCupom("D", "$", "0,00")
        '    BemaRetorno = Daruma_FI_EfetuaFormaPagamentoDescricaoForma("Dinheiro", xString, "")
        '    BemaRetorno = Daruma_FI_TerminaFechamentoCupom("Fechado Automaticamente pelo Sistema")
        'End If
    ElseIf lImpDarumaFW Then
        ' AQUI ESTUDAR COMO SABER SE UM CUPOM ESTÁ ABERTO E COMO FECHA-LO
'        xString = Space(2)
'        BemaRetorno = Daruma_FI_StatusCupomFiscal(xString)
'        If Mid(xString, 1, 1) = "1" Then
'            DarumaBuscaRetorno
'            xString = Space(18)
'            BemaRetorno = Daruma_FI_SaldoAPagar(xString)
'            BemaRetorno = Daruma_FI_IniciaFechamentoCupom("D", "$", "0,00")
'            BemaRetorno = Daruma_FI_EfetuaFormaPagamentoDescricaoForma("Dinheiro", xString, "")
'            BemaRetorno = Daruma_FI_TerminaFechamentoCupom("Fechado Automaticamente pelo Sistema")
'        End If
    End If
    
    
    If ConfiguracaoDiversa.LocalizarCodigo(1, "PETROMOVELAUTO AUTORIZA NFCE") Then
       If ConfiguracaoDiversa.Verdadeiro Then
          VerificarAtivarPetromovelAuto (True)
       End If
    End If
    
    'MsgBox "TESTE ALEX 9"
   
    lLinhasEntreCV = 2
    If ConfiguracaoDiversa.LocalizarCodigo(1, "ECF: PULAR X LINHAS ENTRE CV") Then
        lLinhasEntreCV = ConfiguracaoDiversa.Codigo
    End If
    
    'MsgBox "TESTE ALEX 10"
    
    lCaixaIndividual = False
    If ConfiguracaoDiversa.LocalizarCodigo(g_empresa, "CAIXA DE PISTA INDIVIDUAL") Then
        lCaixaIndividual = ConfiguracaoDiversa.Verdadeiro
    End If
    
    'MsgBox "TESTE ALEX 11"

    
    lGeraCaixaDinheiro = False
    If ConfiguracaoDiversa.LocalizarCodigo(1, "ECF: GERA CAIXA DINHEIRO") Then
        lGeraCaixaDinheiro = ConfiguracaoDiversa.Verdadeiro
    End If

    'MsgBox "TESTE ALEX 12"

    lGeraCaixaChequeAVista = False
    If ConfiguracaoDiversa.LocalizarCodigo(1, "ECF: GERA CAIXA CHEQUE A VISTA") Then
        lGeraCaixaChequeAVista = ConfiguracaoDiversa.Verdadeiro
    End If

    lUtilizaPrecoAPrazo = False
    If ConfiguracaoDiversa.LocalizarCodigo(1, "AUTOMACAO:UTILIZA PRECO A PRAZO") Then
       lUtilizaPrecoAPrazo = ConfiguracaoDiversa.Verdadeiro
    End If
    
'    If lCaixaIndividual = True Then
'        mnRelatorio.Enabled = True
'    Else
'        mnRelatorio.Enabled = False
'    End If


   ' MsgBox "TESTE ALEX 13"

    AutomacaoInicio
    lOrigemAutomacao = False
    lGrupoCombustivel = 4
    lCodigoCartao = 0
    lIlha = 1
    
    lPortaRfid = 0
    lPortaRfid = Val(ReadINI("CUPOM FISCAL", "Porta Leitor RFID", gArquivoIni))
    
    
    
    
    'MsgBox "Form_Load FINALIZADO- TESTE ALEX"
    
    
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If lFinalizaAutomatico = False Then
        If l_flag_cupom_fiscal = "A" Then
            MsgBox "Finalize a NFC-e que está aberta!", vbInformation, "Operação Não Aceita!"
            Cancel = True
            If cmdIniciaProcessoFinalizacaoNFCe.Visible Then
                cmdIniciaProcessoFinalizacaoNFCe.SetFocus
                Exit Sub
            End If
        End If
        If (MsgBox("Deseja realmente sair da NFC-e?", 4 + 32 + 256, "Sair da NFC-e!")) = 7 Then
            Cancel = True
        End If
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CriaLogCupom("Cupom Fiscal Automação: Finalizado.")
    Finaliza
End Sub

Private Sub lbl_automacao_valor_Click(Index As Integer)
    Dim xValor As Currency
    
    If g_nome_empresa Like "*SARA*" Or g_nome_empresa Like "*CENTRAL*" Or g_nome_empresa Like "*P. G.II*" Or g_nome_empresa Like "*GRANADA*" Then
        If lAutomacaoStatusBico(Index) = 6 And ValidaCliente Then
            If (MsgBox("Deseja realmente receber este abastecimento?", vbQuestion + vbYesNo + vbDefaultButton2, "Recebimento de Abastecimento")) = vbYes Then
                xValor = fValidaValor(lbl_automacao_valor(Index).Caption)
                lAutomacaoBicoEmAcerto = lAutomacaoBico(Index)
                lAutomacaoFlagVendaAutomatica = True
                lAutomacaoBicoEmAcerto = lAutomacaoBico(Index)
                lAutomacaoDataEmAcerto = lAutomacaoData(Index)
                lAutomacaoHoraEmAcerto = lAutomacaoHora(Index)
                lAutomacaoTempoAbastecimentoEmAcerto = "" & lAutomacaoTempoAbastecimento(Index)
                lAutomacaoFormaImpressao = "BOTAO"
                
                'Atribui Abastecimento como concluído com numero de ECF = 0
                Call AutomacaoAlteraTabeAbastecimento(Index, 0)
                
                lAutomacaoBicoEmAcerto = 0
                lAutomacaoDataEmAcerto = 0
                lAutomacaoHoraEmAcerto = 0
                lAutomacaoTempoAbastecimentoEmAcerto = ""
                lAutomacaoFormaImpressao = ""
                lAutomacaoBicoEmAcerto = 0
                lAutomacaoFlagVendaAutomatica = False
                lAutomacaoStatusBico(Index) = 0
                lbl_automacao_valor(Index).Caption = ""
                lAutomacaoCodigoProduto(Index) = 0
                lAutomacaoBico(Index) = 0
                lAutomacaoData(Index) = 0
                lAutomacaoHora(Index) = 0
                lAutomacaoTempoAbastecimento(Index) = ""
                lAutomacaoValorLitro(Index) = 0
                lAutomacaoLitros(Index) = 0
                lAutomacaoTotalAPagar(Index) = 0
                'Chama Caixa de Pista
                If (MsgBox("Este recebimento foi feito em dinheiro?", vbQuestion + vbYesNo + vbDefaultButton2, "Tipo de Recebimento!")) = vbYes Then
                    If IntegracaoCaixa.LocalizarNome(g_empresa, "DINHEIRO") Then
                        MovCaixaPista.Empresa = g_empresa
                        MovCaixaPista.Data = lData
                        MovCaixaPista.NumeroMovimento = 1
                        MovCaixaPista.Valor = xValor
                        MovCaixaPista.CodigoHistorico = IntegracaoCaixa.HistoricoPadrao
                        MovCaixaPista.DadosInterno = "CX.PISTA DIV.|@|"
                        MovCaixaPista.CodigoLancamentoPadrao = 8
                        MovCaixaPista.NumeroDocumento = ""
                        MovCaixaPista.Complemento = ""
                        MovCaixaPista.NumeroContaDebito = IntegracaoCaixa.ContaDebito
                        MovCaixaPista.NumeroContaCredito = IntegracaoCaixa.ContaCredito
                        MovCaixaPista.CodigoUsuario = g_usuario
                        MovCaixaPista.TipoMovimento = lTipoMovimento
                        MovCaixaPista.Periodo = lPeriodo
                        MovCaixaPista.NumeroIlha = lIlha
                        MovCaixaPista.DataDigitacao = Format(Now, "dd/mm/yyyy")
                        MovCaixaPista.HoraDigitacao = Format(Now, "HH:mm:ss")
                        MovCaixaPista.DataAlteracao = "00:00:00"
                        MovCaixaPista.HoraAlteracao = "00:00:00"
                        If Not MovCaixaPista.Incluir Then
                            MsgBox "Erro ao incluir movimento no caixa de pista.", vbCritical, "Erro de Integridade!"
                        End If
                    Else
                        MsgBox "Não existe a integração=" & "Dinheiro" & ".", vbCritical, "Erro de Integridade!"
                    End If
                Else
                    mnuCaixaPista_Click
                End If
            End If
        End If
    End If
End Sub
Private Sub lbl_automacao_valor_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Shift = 2 And ValidaCliente Then
        If UCase(g_cidade_empresa) Like "*REDEN*" Or UCase(g_cidade_empresa) Like "*CUMAR*" Or UCase(g_cidade_empresa) Like "*CONCEI*" Then
            If fValidaValor(lbl_automacao_valor(Index).Caption) > 0 Then
                Call GeraDescontoAbastecimento(lAutomacaoBico(Index), lAutomacaoData(Index), lAutomacaoHora(Index), lAutomacaoValorLitro(Index), lAutomacaoTotalAPagar(Index))
            End If
        End If
    End If
End Sub
Private Sub lbl_fila_total_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Shift = 2 Then
        If UCase(g_cidade_empresa) Like "*REDEN*" Or UCase(g_cidade_empresa) Like "*CUMAR*" Or UCase(g_cidade_empresa) Like "*CONCEI*" Then
            If fValidaValor(MovimentoAbastecimento.ValorTotal) > 0 Then
                Call GeraDescontoAbastecimento(MovimentoAbastecimento.Bico, MovimentoAbastecimento.Data, MovimentoAbastecimento.Hora, MovimentoAbastecimento.ValorUnitario, MovimentoAbastecimento.ValorTotal)
                cmd_fila_sair_Click
            End If
        End If
    End If
End Sub

Private Sub mnAbastecimentoAutomacaoFuncionario_Click()
    Dim xChamaCaixa As Boolean
    
    xChamaCaixa = False
    Call GravaAuditoria(1, Me.name, 23, mnAbastecimentoAutomacaoFuncionario.Caption & " Func.:" & l_nome_funcionario)
    If lCaixaIndividual Then
       gStringChamada = "Movimento NFCe Auto" & "|@|" & Format(lDataCupom, "dd/mm/yyyy") & "|@|" & AberturaCaixa.Periodo & "|@|" & l_codigo_funcionario & "|@|"
       Call menu_personalizado.GravaSgpCadastroIni("EmissaoAbastecimentoAutomacaoFuncionario")
    Else
        MsgBox "Funcionalidade disponível apenas para caixa individual", vbInformation + vbExclamation, "Acesso negado"
    End If

End Sub

Private Sub mnFaltaCaixa_Click()
    Call GravaAuditoria(1, Me.name, 23, mnFaltaCaixa.Caption & " Func.:" & l_nome_funcionario)
    If lCaixaIndividual Then
       Dim xChamada As String
       xChamada = "Movimento NFCe Auto" & "|@|" & Format(lDataCupom, "dd/mm/yyyy") & "|@|" & AberturaCaixa.Periodo & "|@|" & l_codigo_funcionario & "|@|"
       Call menu_personalizado.ChamaProgramaVB6ClickMenu("RE", "Falta de Caixa", xChamada)
    Else
        MsgBox "Funcionalidade disponível apenas para caixa individual", vbInformation + vbExclamation, "Acesso negado"
    End If
End Sub

Private Sub mnuCaixaPista_Click()
    Dim xChamaCaixa As Boolean
    
    xChamaCaixa = False
    BuscaPeriodo
    Call GravaAuditoria(1, Me.name, 23, mnuCaixaPista.Caption & " Func.:" & l_nome_funcionario)
    If lCaixaIndividual Then
        If Not AberturaCaixa.LocalizarUltAbertoDataFunc(g_empresa, lDataCupom, "NF", 1, lTipoMovimento, l_codigo_funcionario) Then
            If (MsgBox("O Caixa não encontra-se aberto!" & Chr(10) & "Deseja abrir agora?", vbQuestion + vbYesNo + vbDefaultButton1, "Abertura de Caixa!")) = vbYes Then
                'gStringChamada = "IncluirCompleto|@|" & msk_data.Text & "|@|" & Val(cbo_periodo.Text) & "|@|" & 1 & "|@|" & l_codigo_funcionario & "|@|"
                Call CriaAberturaCaixa(lTipoMovimento, lPeriodo)
                'Call menu_personalizado.GravaSgpCadastroIni("MovimentoAberturaCaixa")
                xChamaCaixa = True
            Else
                MsgBox "O Caixa atual não foi aberto!" & Chr(10) & "Não será possível acessar o caixa sem antes abri-lo?", vbInformation + vbExclamation, "Caixa Inexistente!"
            End If
        Else
            xChamaCaixa = True
        End If
    Else
    
'        If lTipoMovimento = 3 Then
'            lPeriodo = PeriodoTrocaOleo.Periodo
'        End If
    
        If Not AberturaCaixa.LocalizarCxData(g_empresa, lDataCupom, "NF", lPeriodo, 1, lTipoMovimento) Then
            If (MsgBox("O Caixa não encontra-se aberto!" & Chr(10) & "Deseja abrir agora?", vbQuestion + vbYesNo + vbDefaultButton1, "Abertura de Caixa!")) = vbYes Then
                'gStringChamada = "IncluirCompleto|@|" & msk_data.Text & "|@|" & Val(cbo_periodo.Text) & "|@|" & 1 & "|@|" & l_codigo_funcionario & "|@|"
                Call CriaAberturaCaixa(lTipoMovimento, lPeriodo)
                'Call menu_personalizado.GravaSgpCadastroIni("MovimentoAberturaCaixa")
                xChamaCaixa = True
            Else
                MsgBox "O Caixa atual não foi aberto!" & Chr(10) & "Não será possível acessar o caixa sem antes abri-lo?", vbInformation + vbExclamation, "Caixa Inexistente!"
            End If
        Else
            xChamaCaixa = True
        End If
        
    End If
    
    If xChamaCaixa Then
        If lCaixaIndividual Then
            gStringChamada = Format(lDataCupom, "dd/mm/yyyy") & "|@|" & AberturaCaixa.Periodo & "|@|" & 2 & "|@|" & 1 & "|@|" & l_codigo_funcionario & "|@|" & "NF" & "|@|"
        Else
            If lTipoMovimento = 3 Then
                gStringChamada = Format(lDataCupom, "dd/mm/yyyy") & "|@|" & lPeriodo & "|@|" & lTipoMovimento & "|@|" & 1 & "|@|" & l_codigo_funcionario & "|@|" & "NF" & "|@|"
            Else
                gStringChamada = Format(lDataCupom, "dd/mm/yyyy") & "|@|" & lPeriodo & "|@|" & 2 & "|@|" & 1 & "|@|" & l_codigo_funcionario & "|@|" & "NF" & "|@|"
            End If
        End If
        'Call CriaLogCupom("Bematech_FI_FechaPortaSerial")
        'BemaRetorno = Bematech_FI_FechaPortaSerial()
        'Call CriaLogCupom("Bematech_FI_FechaPortaSerial - BemaRetorno=" & BemaRetorno)
        'Call menu_personalizado.GravaSgpCadastroIni("MovimentoCaixaPista")
        Call menu_personalizado.GravaSgpNetCadastroIni("MovimentoCaixaPista")
    End If
End Sub
Private Sub mnuCalculadora_Click()
    Dim retval As Long
    Call GravaAuditoria(1, Me.name, 23, mnuCalculadora.Caption & " Func.:" & l_nome_funcionario)
    retval = Shell("calc", vbNormalFocus)
End Sub
Private Sub mnuCancelaCartao_Click()
    Dim xNomeCartao As String
    
    If l_flag_cupom_fiscal = "A" Then
        MsgBox "Esta função não poderá ser executada com cupom aberto." & vbCrLf & "Feche o cupom que encontra-se aberto.", vbInformation + vbOKOnly, "Cupom Fiscal Aberto!"
        Exit Sub
    End If
    Call GravaAuditoria(1, Me.name, 23, mnuCancelaCartao.Caption & " Func.:" & l_nome_funcionario)
    Call AtivaDesativaTimer(False)
    gNumeroControleSolicitacao = Configuracao.ProximaSolicitacaoTEF(g_empresa)
    'Set CerradoTef = Nothing
    Set CerradoTef = New CerradoComponenteTef
    g_string = ""
    frm_tipo_cartao.Show 1
    xNomeCartao = g_string
    g_string = ""
    lRespostaTEF = CerradoTef.SolicitacaoCNC("NFCe", gNumeroControleSolicitacao, gQtdViasTEF, xNomeCartao, l_codigo_funcionario, l_nome_funcionario)
    Set CerradoTef = Nothing
    Call AtivaDesativaTimer(True)
    mnuSenha_Click
End Sub

Private Sub mnuCancelamentoNFCe_Click()
    If VerificaRestricaoCancelamento Then
        Call menu_personalizado.GravaSgpNetCadastroIni("CancelaNFCe")
    End If
End Sub
Private Sub mnuReimpressaoNFCe_Click()
    Call menu_personalizado.GravaSgpNetCadastroIni("REIMPRESSAONFCe")
End Sub
Function VerificaRestricaoCancelamento() As Boolean
    VerificaRestricaoCancelamento = True
    
    If ConfiguracaoDiversa.LocalizarCodigo(1, "RESTRINGE:CANCELAMENTO DE NFCe") Then
        If ConfiguracaoDiversa.Codigo <= g_nivel_acesso Then
            MsgBox "Este usuário não está autorizado a realizar esta operação", vbInformation, "Operação Não Autorizada!"
            VerificaRestricaoCancelamento = False
        End If
    End If
End Function


Private Sub mnuConsultaCheque_Click()
    Call GravaAuditoria(1, Me.name, 23, mnuConsultaCheque.Caption & " Func.:" & l_nome_funcionario)
    Call menu_personalizado.GravaCheqPostoIni("consultaCheq")
'    Dim xCpfCnpj As String
'    Dim xCpfCnpjMasc As String
'    Dim xMensagem As String
'
'    xCpfCnpj = Trim(InputBox("Digite o CPF ou CNPJ", "Consulta de Cheque"))
'    If xCpfCnpj <> "" Then
'        If Len(xCpfCnpj) = 11 Then
'            xCpfCnpjMasc = fMascaraCPF(xCpfCnpj)
'            xMensagem = "CPF " & xCpfCnpjMasc & Chr(10)
'            If CalculaDigitoCPF(xCpfCnpj) = False Then
'                'MsgBox "Favor verificar o " & xMensagem & " informado.", vbCritical, "Erro na rotina de dígito!"
'                Exit Sub
'            End If
'        ElseIf Len(xCpfCnpj) = 14 Then
'            xCpfCnpjMasc = fMascaraCNPJ(xCpfCnpj)
'            xMensagem = "CNPJ " & xCpfCnpjMasc & Chr(10)
'            If CalculaDigitoCNPJ(xCpfCnpj) = False Then
'                'MsgBox "Favor verificar o " & xMensagem & " informado.", vbCritical, "Erro na rotina de dígito!"
'                Exit Sub
'            End If
'        End If
'        If Len(xCpfCnpj) = 11 Or Len(xCpfCnpj) = 14 Then
'            g_string = MovimentoChequeDevolvido.LocalizarCpfCnpj(xCpfCnpj)
'            If g_string <> "" Then
'                MsgBox xMensagem & "Este cliente tem " & RetiraGString(2) & " cheque(s) devoldido(s)." & Chr(10) & "Valor do débito R$ " & Format(CCur(RetiraGString(1)), "###,###,##0.00"), vbCritical, "Cliente em Débido!"
'                g_string = ""
'            Else
'                If MovimentoCheque.LocalizarCpfCnpj(xCpfCnpj) <> "" Then
'                    MsgBox xMensagem & "Este é um cliente cadastrado.", vbInformation, "Nada Consta!"
'                Else
'                    MsgBox xMensagem & "Este não é um cliente cadastrado!", vbInformation, "Nada Consta!"
'                End If
'            End If
'        Else
'            MsgBox "O número informado não é um CPF ou CNPJ válido!" & Chr(10) & Chr(10) & "CPF deve ter 11 números." & Chr(10) & "CNPJ deve ter 14 números." & Chr(10) & Chr(10) & "O número informado foi: " & xCpfCnpj & " com " & Len(xCpfCnpj) & " dígitos", vbInformation, "Número Inválido!"
'        End If
'    End If
End Sub
Private Sub mnuFechamentoCaixa_Click()
    Dim i As Integer
    Dim retval As Long
    Dim xContinua As Boolean
    Dim xTipoVenda As String
    Dim xMensagem As String
    Dim xExcluiCartoesIdentFid As Boolean
    
    If l_flag_cupom_fiscal = "A" Then
        MsgBox "Esta função não poderá ser executada com cupom aberto." & vbCrLf & "Feche o cupom que encontra-se aberto.", vbInformation + vbOKOnly, "Cupom Fiscal Aberto!"
        Exit Sub
    End If

    BuscaPeriodo
    xExcluiCartoesIdentFid = False
'    Código comentado pelo fato da necessidade de continuar para
'    ter LOG dos ENCERRANTES
'    If MovimentoBomba.LocalizarCodigo(g_empresa, g_cfg_data_i, Val(g_cfg_periodo_i), 1, 999) Then
'        MsgBox "Já existe movimento de bomba lançado neste caixa.", vbInformation, "Operação não aceita!"
'        Exit Sub
'    End If
    If (MsgBox("Deseja fazer leitura dos encerrantes?", vbDefaultButton2 + vbYesNo + vbQuestion, "Encerrante Automático")) = vbYes Then
        xMensagem = "Empresa: " & g_nome_empresa & vbCrLf
        xMensagem = xMensagem & "Data: " & Format(Date, "dd/mm/yyyy") & " as " & Format(Time, "HH:MM:SS") & vbCrLf & vbCrLf
        xMensagem = xMensagem & "Foi pedido o fechamento do caixa:" & g_cfg_periodo_i & " da Data:" & Format(g_cfg_data_i, "dd/mm/yyyy") & vbCrLf
        xMensagem = xMensagem & "Pelo funcionário:" & l_nome_funcionario & vbCrLf
        'Call EnviaMensagemEmail(g_empresa, g_nome_empresa, "Fechamento Caixa!", xMensagem, False, 0)
        Call GravaAuditoria(1, Me.name, 23, mnuFechamentoCaixa.Caption & " Func.:" & l_nome_funcionario)
        If g_automacao Then
            Call TelaAguarde("Aguarde! Lendo encerrantes das bombas...", True)
            TimerAutomacao.Enabled = False
            'g_string = "Automacao|@|" & l_codigo_funcionario & "|@|" & l_nome_funcionario & "|@|" & g_cfg_data_i & "|@|" & g_cfg_periodo_i & "|@|"
            'movimento_bomba.Show 1
            xContinua = False
            If lMarcaAutomacao = "COMPANY" Then
                If ComunicaAutomacaoCerradoBD("AUTOMACAO LEITURA ENCERRANTE", g_cfg_data_i & "|@|" & g_cfg_periodo_i & "|@|") Then
                    xContinua = True
                Else
                    Call TelaAguarde("", False)
                    MsgBox "Não foi possível comunicar com o progama AutoCerradoCompany.", vbCritical, "Erro de Comunicação com AutoCerradoCompany!"
                    Call GravaAuditoria(1, Me.name, 28, "Não foi possível comunicar com o AutoCerradoCompany")
                    MsgBox "Após a confirmação desta mensagem, o sistema irá abrir automaticamento o progama AutoCerradoCompany." & vbCrLf & "Após este procedimento, favor tentar o fechamento novamente.", vbCritical, "Mensagem Importante sobre Automação!"
                    LoopAbreAutomacao
                End If
            ElseIf lMarcaAutomacao = "HOROUSTECH" Then
                If ComunicaAutomacaoCerradoBD("AUTOMACAO LEITURA ENCERRANTE", g_cfg_data_i & "|@|" & g_cfg_periodo_i & "|@|") Then
                    xContinua = True
                Else
                    Call TelaAguarde("", False)
                    MsgBox "Não foi possível comunicar com o progama AutoCerradoHorousTech.", vbCritical, "Erro de Comunicação com AutoCerradoHorousTech!"
                    Call GravaAuditoria(1, Me.name, 28, "Não foi possível comunicar com o AutoCerradoHorousTech")
                    MsgBox "Após a confirmação desta mensagem, o sistema irá abrir automaticamento o progama AutoCerradoHorousTech." & vbCrLf & "Após este procedimento, favor tentar o fechamento novamente.", vbCritical, "Mensagem Importante sobre Automação!"
                    LoopAbreAutomacao
                End If
            ElseIf lMarcaAutomacao = "EZTECH" Then
                If ComunicaAutomacaoCerradoBD("AUTOMACAO LEITURA ENCERRANTE", g_cfg_data_i & "|@|" & g_cfg_periodo_i & "|@|") Then
                    Call GravaLogEncerrantes
                    xContinua = True
                Else
                    Call TelaAguarde("", False)
                    MsgBox "Não foi possível comunicar com o progama AutoCerradoEZ.", vbCritical, "Erro de Automação!"
                End If
            ElseIf lMarcaAutomacao = "IONICS" Then
                If ComunicaAutomacaoIonics("AUTOMACAO LEITURA ENCERRANTE", g_cfg_data_i & "|@|" & g_cfg_periodo_i & "|@|") Then
                    Call GravaLogEncerrantes
                    xContinua = True
                Else
                    Call TelaAguarde("", False)
                    MsgBox "Não foi possível comunicar com o progama AutoCerradoIonics.", vbCritical, "Erro de Automação!"
                End If
            End If
            If xContinua Then
                Call AlteraMovBombaParaSubCaixa(g_cfg_data_i, g_cfg_periodo_i)
                If lMarcaAutomacao = "COMPANY" Then
                    Call GravaLogEncerrantes
                    'retval = Shell("C:\Cerrado\AutoCerradoCompany\AutoCerradoCompany.exe", vbNormalFocus)
                ElseIf lMarcaAutomacao = "HOROUSTECH" Then
                    Call GravaLogEncerrantes
                    'retval = Shell("C:\Cerrado.Net\AutoCerradoHorousTech\AutoCerradoHorousTech.exe", vbNormalFocus)
                End If
                For i = 1 To lQtdBomba
                    If Bomba.LocalizarCodigo(g_empresa, i) Then
                        If EncerranteAtual.LocalizarCodigo(g_empresa, i) Then
                            MovimentoBomba.Empresa = g_empresa
                            MovimentoBomba.Data = g_cfg_data_i 'lData
                            MovimentoBomba.Periodo = g_cfg_periodo_i 'lPeriodo
                            MovimentoBomba.SubCaixa = 999
                            MovimentoBomba.CodigoBomba = Bomba.Codigo
                            MovimentoBomba.Abertura = MovimentoBomba.EncerranteBicoAnatesDataPeriodo(g_empresa, lData, i, 9, 999)
                            MovimentoBomba.Encerrante = EncerranteAtual.Encerrante
                            MovimentoBomba.QuantidadeSaida = MovimentoBomba.Encerrante - MovimentoBomba.Abertura
                            MovimentoBomba.PrecoCusto = Bomba.PrecoCusto
                            MovimentoBomba.PrecoVenda = Bomba.PrecoVenda
                            MovimentoBomba.TipoCombustivel = Bomba.TipoCombustivel
                            MovimentoBomba.NumeroTanque = Bomba.NumeroTanque
                            MovimentoBomba.NumeroIlha = lIlha
                            MovimentoBomba.TotalVendaBruto = MovimentoBomba.QuantidadeSaida * MovimentoBomba.PrecoVenda
                            
                            Call CalculaTotalDescontoAcrescimo(MovimentoBomba.Data, MovimentoBomba.Periodo, MovimentoBomba.TipoCombustivel, MovimentoBomba.CodigoBomba, MovimentoBomba.SubCaixa)
                            
                            MovimentoBomba.TotalDesconto = lTotalDescontoNivelPreco
                            MovimentoBomba.TotalAcrescimo = lTotalAcrescimoNivelPreco
                            
                            If MovimentoBomba.Incluir Then
                                MovimentoBombaEscritorio.Empresa = g_empresa
                                MovimentoBombaEscritorio.Data = g_cfg_data_i 'lData
                                MovimentoBombaEscritorio.Periodo = g_cfg_periodo_i 'lPeriodo
                                MovimentoBombaEscritorio.SubCaixa = 999
                                MovimentoBombaEscritorio.CodigoBomba = Bomba.Codigo
                                MovimentoBombaEscritorio.Abertura = MovimentoBombaEscritorio.EncerranteBicoAnatesDataPeriodo(g_empresa, lData, i, 9, 999)
                                MovimentoBombaEscritorio.Encerrante = EncerranteAtual.Encerrante
                                MovimentoBombaEscritorio.QuantidadeSaida = MovimentoBombaEscritorio.Encerrante - MovimentoBombaEscritorio.Abertura
                                MovimentoBombaEscritorio.PrecoCusto = Bomba.PrecoCusto
                                MovimentoBombaEscritorio.PrecoVenda = Bomba.PrecoVenda
                                MovimentoBombaEscritorio.TipoCombustivel = Bomba.TipoCombustivel
                                MovimentoBombaEscritorio.NumeroTanque = Bomba.NumeroTanque
                                MovimentoBombaEscritorio.NumeroIlha = lIlha
                                MovimentoBombaEscritorio.TotalVendaBruto = MovimentoBombaEscritorio.QuantidadeSaida * MovimentoBombaEscritorio.PrecoVenda
                                MovimentoBombaEscritorio.TotalDesconto = lTotalDescontoNivelPreco
                                MovimentoBombaEscritorio.TotalAcrescimo = lTotalAcrescimoNivelPreco
                                
                                If Not MovimentoBombaEscritorio.Incluir Then
                                    MsgBox "Não foi possível incluir o movimento de bomba-esc. do bico=" & i, vbInformation, "Duplicidade de Registro"
                                Else
                                    CriaLogECF ("[LOG DE TESTE AGORA] - PASSO 01")
                                End If
                            Else
                                MsgBox "Não foi possível incluir o movimento de bomba do bico=" & i, vbInformation, "Duplicidade de Registro"
                            End If
                        End If
                    End If
                Next
                If lData < g_cfg_data_i Then
                    lData = g_cfg_data_i
                End If
                CriaLogECF ("[LOG DE TESTE AGORA] - PASSO 02")
                LoopIncluiMovBombaCaixa
                CriaLogECF ("[LOG DE TESTE AGORA] - PASSO 04")
                Call TelaAguarde("", False)
                'Call CriaLogCupom("Cupom Fiscal: A Emissão do Cupom Complementar Foi Chamada Automaticamente.")
                'emissao_cupom_complementar.Show 1
                If g_automacao = True Then
                    Call CriaLogCupom("NFCe Auto: A Emissão da NFCe Complementar Automação Foi Chamada Automaticamente.")
                    EmissaoNFCeComplementarAuto.Show 1
                Else
                    Call CriaLogCupom("NFCe Auto: A Emissão da NFCe Complementar Foi Chamada Automaticamente.")
                    EmissaoNFCeComplementar.Show 1
                End If
                Call CriaLogCupom("Cupom Fiscal: Foi Retornado ao Cupom Fiscal com o Retorno: " & Chr(39) & RetiraGString(1) & Chr(39))
                'If RetiraGString(1) = "imprimiu" Then
                    g_string = ""
                    'If (MsgBox("Deseja realmente mudar o período do caixa atual para o próximo periodo de caixa?", vbDefaultButton2 + vbYesNo + vbQuestion, "Fechamento de Caixa")) = vbYes Then
                    MsgBox "O sistema irá mudar o período do caixa atual para o próximo periodo de caixa", vbOKOnly + vbInformation, "Fechamento de Caixa"
                        Call GravaAuditoria(1, Me.name, 26, "Foi confirmado a mudança de período:" & lPeriodo)
                        xTipoVenda = ReadINI("CUPOM FISCAL", "Tipo de Venda", gArquivoIni)
                        If xTipoVenda = "AUTOMACAO/CONVENIENCIA" Then
                            If LiberacaoDigitacao.LocalizarCodigo(g_empresa, 2) Then
                                If LiberacaoDigitacao.PeriodoInicial = gQtdPeriodo Then
                                    LiberacaoDigitacao.DataInicial = LiberacaoDigitacao.DataInicial + 1
                                    LiberacaoDigitacao.DataFinal = LiberacaoDigitacao.DataFinal + 1
                                    LiberacaoDigitacao.PeriodoInicial = 0
                                    LiberacaoDigitacao.PeriodoFinal = 0
                                    If lMarcaAutomacao = "COMPANY" Then
                                        'If g_nome_empresa Like "*VERA CRUZ*" Or g_nome_empresa Like "*BRITO BARROS*" Or g_nome_empresa Like "*AUTO POSTO MT*" Then
                                        'If g_nome_empresa Like "*BRITO BARROS*" Or g_nome_empresa Like "*AUTO POSTO MT*" Then
                                        If g_nome_empresa Like "*BRITO BARROS*" Then
                                            xExcluiCartoesIdentFid = True
                                        End If
                                    ElseIf lMarcaAutomacao = "HOROUSTECH" Then
                                        'AM SAO JUDAS
'                                        If g_nome_empresa Like "*Posto Jd. Helvecia*" Then
'                                            xExcluiCartoesIdentFid = True
'                                        End If
                                    End If
                                End If
                                LiberacaoDigitacao.PeriodoInicial = LiberacaoDigitacao.PeriodoInicial + 1
                                LiberacaoDigitacao.PeriodoFinal = LiberacaoDigitacao.PeriodoFinal + 1
                                If LiberacaoDigitacao.Alterar(g_empresa, 2) Then
                                    If LiberacaoDigitacao.LocalizarCodigo(g_empresa, 3) Then
                                        If LiberacaoDigitacao.PeriodoInicial = gQtdPeriodo Then
                                            LiberacaoDigitacao.DataInicial = LiberacaoDigitacao.DataInicial + 1
                                            LiberacaoDigitacao.DataFinal = LiberacaoDigitacao.DataFinal + 1
                                            LiberacaoDigitacao.PeriodoInicial = 0
                                            LiberacaoDigitacao.PeriodoFinal = 0
                                        End If
                                        LiberacaoDigitacao.PeriodoInicial = LiberacaoDigitacao.PeriodoInicial + 1
                                        LiberacaoDigitacao.PeriodoFinal = LiberacaoDigitacao.PeriodoFinal + 1
                                        If Not LiberacaoDigitacao.Alterar(g_empresa, 3) Then
                                            MsgBox "Erro ao alterar a Liberação de Digitação!", vbInformation, "Registro não Alterado!"
                                        End If
                                    End If
                                Else
                                    MsgBox "Erro ao alterar a Liberação de Digitação!", vbInformation, "Registro não Alterado!"
                                End If
                            End If
                        Else
                            If LiberacaoDigitacao.LocalizarCodigo(g_empresa, 2) Then
                                ' Se empresa = POSTO PELICANO II
                                If g_nome_empresa Like "*POSTO PELICANO II*" Or g_nome_empresa Like "*TEIXEIRA E PINHEIRO LTDA*" Then
                                    'If LiberacaoDigitacao.PeriodoInicial = 2 Then
                                    If LiberacaoDigitacao.PeriodoInicial = 3 Then
                                        Dim xDiaSemana As Integer
                                        xDiaSemana = Weekday(LiberacaoDigitacao.DataInicial) '1-Domingo, 2-Segunda ... 6-Sexta, 7-Sabado
                                        'MsgBox "Dia da semana =" & xDiaSemana
                                        ' Segunda a Quinta, vai ter 2 periodos
                                        If xDiaSemana >= 2 And xDiaSemana <= 5 Then
                                            'Call CriaLogCupom("Cupom Fiscal: Foi mudado Periodo 2 para 3 porque dia da semana é: " & xDiaSemana)
                                            Call CriaLogCupom("NFCe: Foi mudado Periodo 3 para 4 porque dia da semana é: " & xDiaSemana)
                                            'MsgBox "dias para 2 periodos"
                                            'LiberacaoDigitacao.PeriodoInicial = 3
                                            LiberacaoDigitacao.PeriodoInicial = 4
                                        End If
                                    End If
                                End If
                                If LiberacaoDigitacao.PeriodoInicial = gQtdPeriodo Then
                                    LiberacaoDigitacao.DataInicial = LiberacaoDigitacao.DataInicial + 1
                                    LiberacaoDigitacao.DataFinal = LiberacaoDigitacao.DataFinal + 1
                                    LiberacaoDigitacao.PeriodoInicial = 0
                                    LiberacaoDigitacao.PeriodoFinal = 0
                                    If lMarcaAutomacao = "COMPANY" Then
                                        'If g_nome_empresa Like "*VERA CRUZ*" Or g_nome_empresa Like "*BRITO BARROS*" Or g_nome_empresa Like "*AUTO POSTO MT*" Then
                                        'If g_nome_empresa Like "*BRITO BARROS*" Or g_nome_empresa Like "*AUTO POSTO MT*" Then
                                        If g_nome_empresa Like "*BRITO BARROS*" Then
                                            xExcluiCartoesIdentFid = True
                                        End If
                                    ElseIf lMarcaAutomacao = "HOROUSTECH" Then
                                        ' AM SAO JUDAS
'                                        If g_nome_empresa Like "*Posto Jd. Helvecia*" Then
'                                            xExcluiCartoesIdentFid = True
'                                        End If
                                    End If
                                End If
                                LiberacaoDigitacao.PeriodoInicial = LiberacaoDigitacao.PeriodoInicial + 1
                                LiberacaoDigitacao.PeriodoFinal = LiberacaoDigitacao.PeriodoFinal + 1
                                If Not LiberacaoDigitacao.Alterar(g_empresa, 2) Then
                                    MsgBox "Erro ao alterar a Liberação de Digitação!", vbInformation, "Registro não Alterado!"
                                End If
                            End If
                        End If
                                            
                        'Aqui começa o processo para zerar os cartoes no identfid.
                        If xExcluiCartoesIdentFid = True Then
                            Call CriaLogAutomacao("Foi solicitado a exclusão dos cartões de abastecimento.")
                            If ComunicaAutomacaoCerradoBD("EXCLUI CARTOES RFID", "") Then
                                Call CriaLogAutomacao("A exclusão dos cartões de abastecimento foi concluída com sucesso.")
                                Dim xQtdCartoesAlterados As Long
                                xQtdCartoesAlterados = Conectar.ExecutaSql("UPDATE CartaoAbastecimento SET [Posicao do Registro] = ''")
                                If xQtdCartoesAlterados > 0 Then
                                    Call CriaLogAutomacao("Foi alterado no banco para excluído " & xQtdCartoesAlterados & " cartões de abastecimento.")
                                Else
                                    Call CriaLogAutomacao("Erro ao alterar no banco para excluído os cartões de abastecimento.")
                                End If
                            Else
                                Call CriaLogAutomacao("Não foi possível excluir os cartões de abastecimento.")
                                MsgBox "Não foi possível comunicar com o progama AutoCerradoCompany." & vbCrLf & "Cartões de Abastecimento não serão excluídos da Automação.", vbCritical, "Erro de Automação, Exclusão Cartões!"
                            End If
                        End If
                                
                        g_cfg_periodo_i = LiberacaoDigitacao.PeriodoInicial
                        g_cfg_periodo_f = LiberacaoDigitacao.PeriodoFinal
                        g_cfg_data_i = LiberacaoDigitacao.DataInicial
                        g_cfg_data_f = LiberacaoDigitacao.DataFinal
                        
                        Me.Caption = "Nota Fiscal do Consumidor Eletrônica - " & l_nome_funcionario & " | Caixa: " & Val(g_cfg_periodo_i) & " Em: " & Format(g_cfg_data_i, "dd/mm/yyyy")
                        
                        If ConfiguracaoDiversa.LocalizarCodigo(1, "PETROMOVELAUTO AUTORIZA NFCE") Then
                           If ConfiguracaoDiversa.Verdadeiro Then
                              Call AtualizaVariaveisGlobaisPetromovelAuto
                           End If
                        End If

                    'Else
                        'Call GravaAuditoria(1, Me.name, 26, "Não foi confirmado o fechamento de caixa:" & lPeriodo)
                    'End If
                'End If
                
                'Busca Retorno das NFCe Complementares Caso Necessário
                If ConfiguracaoDiversa.LocalizarCodigo(1, "NFCe Complementar Sem Retorno") Then
                    If ConfiguracaoDiversa.Verdadeiro = True Then
                        Call TelaAguarde("Aguarde! Processando retorno de NFCe Complementar...", True)
                        Call ConsisteNFCePendentes(0)
                        Call AguardaMS(10000)
                        Call TelaAguarde("", False)
                    End If
                End If
                
                If txt_cliente.Enabled Then
                    NovoCupom
                End If
                TimerAutomacao.Enabled = True
                If txt_cliente.Enabled Then
                    txt_cliente.SetFocus
                Else
                    txt_funcionario_ponto.SetFocus
                End If
            End If
'            Call WriteINI("EMAIL", "Numero do Email", "0", lNomeArquivoAutomacaoIni)
'            Call WriteINI("EMAIL", "Concluido", "SIM", lNomeArquivoAutomacaoIni)
        Else
            g_string = "FechamentoCaixa|@|"
            movimento_bomba.Show 1
            g_string = ""
        End If
    End If
    AtualizaBombasAbastecimento
    AutomacaoMostraBicos
End Sub
Private Sub mnuFuncaoADM_Click()
    Dim xString As String
    
    If l_flag_cupom_fiscal = "A" Then
        MsgBox "Esta função não poderá ser executada com cupom aberto." & vbCrLf & "Feche o cupom que encontra-se aberto.", vbInformation + vbOKOnly, "Cupom Fiscal Aberto!"
        Exit Sub
    End If
    Call GravaAuditoria(1, Me.name, 23, mnuFuncaoADM.Caption & " Func.:" & l_nome_funcionario)
    Call AtivaDesativaTimer(False)
    gNumeroControleSolicitacao = Configuracao.ProximaSolicitacaoTEF(g_empresa)
    'Set CerradoTef = Nothing
    Set CerradoTef = New CerradoComponenteTef
    
    
    
    xString = "Selecione a Bandeira do Cartão Desejado|@|"
    xString = xString & "10|@|"
    xString = xString & "1|@|TecBan|@|"
    xString = xString & "2|@|Ticket Car Smart|@|"
    xString = xString & "3|@|Outros(Visa/Redecard)|@|"
    xString = xString & "4|@|Smart Shop / Check Check|@|"
    xString = xString & "5|@|SuperCard|@|"
    xString = xString & "6|@|HiperCard|@|"
    xString = xString & "7|@|PagCard|@|"
    xString = xString & "8|@|Usa Card|@|"
    xString = xString & "9|@|GodCard|@|"
    xString = xString & "10|@|Cerrado Tef|@|"
    Do Until Len(g_string) = 3
        g_string = xString
        opcaoGeral.Show 1
        If Len(g_string) > 0 Then
            xString = RetiraGString(1)
            Exit Do
        End If
    Loop
    
    lRespostaTEF = False
    Select Case xString
        Case "1"
            lRespostaTEF = CerradoTef.SolicitacaoADM("NFCe", gNumeroControleSolicitacao, gQtdViasTEF, "TecBan", lLinhasEntreCV, l_codigo_funcionario, l_nome_funcionario)
        Case "2"
            lRespostaTEF = CerradoTef.SolicitacaoADM("NFCe", gNumeroControleSolicitacao, gQtdViasTEF, "TCSMART", lLinhasEntreCV, l_codigo_funcionario, l_nome_funcionario)
        Case "3"
            lRespostaTEF = CerradoTef.SolicitacaoADM("NFCe", gNumeroControleSolicitacao, gQtdViasTEF, "Outras", lLinhasEntreCV, l_codigo_funcionario, l_nome_funcionario)
        Case "4"
            lRespostaTEF = CerradoTef.SolicitacaoADM("NFCe", gNumeroControleSolicitacao, gQtdViasTEF, "SMARTEF", lLinhasEntreCV, l_codigo_funcionario, l_nome_funcionario)
        Case "5"
            lRespostaTEF = CerradoTef.SolicitacaoADM("NFCe", gNumeroControleSolicitacao, gQtdViasTEF, "SUPERTEF", lLinhasEntreCV, l_codigo_funcionario, l_nome_funcionario)
        Case "6"
            lRespostaTEF = CerradoTef.SolicitacaoADM("NFCe", gNumeroControleSolicitacao, gQtdViasTEF, "HIPERTEF", lLinhasEntreCV, l_codigo_funcionario, l_nome_funcionario)
        Case "7"
            lRespostaTEF = CerradoTef.SolicitacaoADM("NFCe", gNumeroControleSolicitacao, gQtdViasTEF, "PAGCARD", lLinhasEntreCV, l_codigo_funcionario, l_nome_funcionario)
        Case "8"
            lRespostaTEF = CerradoTef.SolicitacaoADM("NFCe", gNumeroControleSolicitacao, gQtdViasTEF, "TEFNEUS", lLinhasEntreCV, l_codigo_funcionario, l_nome_funcionario)
        Case "9"
            lRespostaTEF = CerradoTef.SolicitacaoADM("NFCe", gNumeroControleSolicitacao, gQtdViasTEF, "GODCARD", lLinhasEntreCV, l_codigo_funcionario, l_nome_funcionario)
        Case "10"
            lRespostaTEF = CerradoTef.SolicitacaoADM("NFCe", gNumeroControleSolicitacao, gQtdViasTEF, "TEFCERRADO", lLinhasEntreCV, l_codigo_funcionario, l_nome_funcionario)
        
'            'teste para fechar gerencial caso esteja aberto
'            If EcfQuickLeRegistrador("EstadoFiscal", "Long", 5) = "32" Then
'                Call EcfQuickEncerraDocumento(0, "Gerencial")
'            End If
'            If EcfQuickLeRegistrador("EstadoFiscal", "Long", 5) = "64" Then
'                Call EcfQuickEncerraDocumento(0, "Gerencial")
'            End If
        
    End Select
    'If MsgBox("Operação administrativas TecBan?" & Chr(10) & Chr(10) & "Sim para TecBan" & Chr(10) & "Não para Outras Bandeiras", vbYesNo + vbDefaultButton2 + vbQuestion, "Operação Administrativas") = vbYes Then
    '    lRespostaTEF = CerradoTef.SolicitacaoADM(gNumeroControleSolicitacao, gQtdViasTEF, "TecBan")
    'Else
    '    lRespostaTEF = CerradoTef.SolicitacaoADM(gNumeroControleSolicitacao, gQtdViasTEF, "Outras")
    'End If
    Set CerradoTef = Nothing
    Call AtivaDesativaTimer(True)
    mnuSenha_Click
End Sub

Private Sub mnuLancamentoEncerrante_Click()
    Dim i As Integer
    Dim retval As Long
    Dim xContinua As Boolean
    'Dim xTipoVenda As String
    Dim xMensagem As String
    
    If l_flag_cupom_fiscal = "A" Then
        MsgBox "Esta função não poderá ser executada com cupom aberto." & vbCrLf & "Feche o cupom que encontra-se aberto.", vbInformation + vbOKOnly, "Cupom Fiscal Aberto!"
        Exit Sub
    End If

    BuscaPeriodo
    If (MsgBox("Deseja fazer leitura dos encerrantes?", vbDefaultButton2 + vbYesNo + vbQuestion, "Encerrante Automático")) = vbYes Then
        xMensagem = "Empresa: " & g_nome_empresa & vbCrLf
        xMensagem = xMensagem & "Data: " & Format(Date, "dd/mm/yyyy") & " as " & Format(Time, "HH:MM:SS") & vbCrLf & vbCrLf
        xMensagem = xMensagem & "Foi pedido o fechamento do caixa:" & g_cfg_periodo_i & " da Data:" & Format(g_cfg_data_i, "dd/mm/yyyy") & vbCrLf
        xMensagem = xMensagem & "Pelo funcionário:" & l_nome_funcionario & vbCrLf
        'Call EnviaMensagemEmail(g_empresa, g_nome_empresa, "Fechamento Caixa!", xMensagem, False, 0)
        Call GravaAuditoria(1, Me.name, 23, "Lançamento do Encerrante Func.:" & l_nome_funcionario)
        Call GravaAuditoria(1, Me.name, 23, "Data: " & Format(g_cfg_data_i, "dd/MM/yyyy") & " Periodo: " & g_cfg_periodo_i)
        Call TelaAguarde("Aguarde! Lendo encerrantes das bombas...", True)
        TimerAutomacao.Enabled = False
        xContinua = False
        If lMarcaAutomacao = "COMPANY" Then
            If ComunicaAutomacaoCerradoBD("AUTOMACAO LEITURA ENCERRANTE", g_cfg_data_i & "|@|" & g_cfg_periodo_i & "|@|") Then
                xContinua = True
            Else
                Call TelaAguarde("", False)
                MsgBox "Não foi possível comunicar com o progama AutoCerradoCompany.", vbCritical, "Erro de Automação!"
            End If
        ElseIf lMarcaAutomacao = "HOROUSTECH" Then
            If ComunicaAutomacaoCerradoBD("AUTOMACAO LEITURA ENCERRANTE", g_cfg_data_i & "|@|" & g_cfg_periodo_i & "|@|") Then
                xContinua = True
            Else
                Call TelaAguarde("", False)
                MsgBox "Não foi possível comunicar com o progama AutoCerradoHorousTech.", vbCritical, "Erro de Automação!"
            End If
        ElseIf lMarcaAutomacao = "EZTECH" Then
            If ComunicaAutomacaoCerradoBD("AUTOMACAO LEITURA ENCERRANTE", g_cfg_data_i & "|@|" & g_cfg_periodo_i & "|@|") Then
                Call GravaLogEncerrantes
                xContinua = True
            Else
                Call TelaAguarde("", False)
                MsgBox "Não foi possível comunicar com o progama AutoCerradoEZ.", vbCritical, "Erro de Automação!"
            End If
        ElseIf lMarcaAutomacao = "IONICS" Then
            If ComunicaAutomacaoIonics("AUTOMACAO LEITURA ENCERRANTE", g_cfg_data_i & "|@|" & g_cfg_periodo_i & "|@|") Then
                Call GravaLogEncerrantes
                xContinua = True
            Else
                Call TelaAguarde("", False)
                MsgBox "Não foi possível comunicar com o progama AutoCerradoIonics.", vbCritical, "Erro de Automação!"
            End If
        End If
        If xContinua Then
            Call AlteraMovBombaParaSubCaixa(g_cfg_data_i, g_cfg_periodo_i)
            If lMarcaAutomacao = "COMPANY" Then
                Call GravaLogEncerrantes
                'retval = Shell("C:\Cerrado\AutoCerradoCompany\AutoCerradoCompany.exe", vbNormalFocus)
            ElseIf lMarcaAutomacao = "HOROUSTECH" Then
                Call GravaLogEncerrantes
                'retval = Shell("C:\Cerrado.Net\AutoCerradoHorousTech\AutoCerradoHorousTech.exe", vbNormalFocus)
            End If
            For i = 1 To lQtdBomba
                If Bomba.LocalizarCodigo(g_empresa, i) Then
                    If EncerranteAtual.LocalizarCodigo(g_empresa, i) Then
                        MovimentoBomba.Empresa = g_empresa
                        MovimentoBomba.Data = g_cfg_data_i 'lData
                        MovimentoBomba.Periodo = g_cfg_periodo_i 'lPeriodo
                        MovimentoBomba.SubCaixa = 999
                        MovimentoBomba.CodigoBomba = Bomba.Codigo
                        MovimentoBomba.Abertura = MovimentoBomba.EncerranteBicoAnatesDataPeriodo(g_empresa, lData, i, 9, 999)
                        MovimentoBomba.Encerrante = EncerranteAtual.Encerrante
                        MovimentoBomba.QuantidadeSaida = MovimentoBomba.Encerrante - MovimentoBomba.Abertura
                        MovimentoBomba.PrecoCusto = Bomba.PrecoCusto
                        MovimentoBomba.PrecoVenda = Bomba.PrecoVenda
                        MovimentoBomba.TipoCombustivel = Bomba.TipoCombustivel
                        MovimentoBomba.NumeroTanque = Bomba.NumeroTanque
                        MovimentoBomba.NumeroIlha = lIlha
                        MovimentoBomba.TotalVendaBruto = MovimentoBomba.QuantidadeSaida * MovimentoBomba.PrecoVenda
                        
                       
                        Call CalculaTotalDescontoAcrescimo(MovimentoBomba.Data, MovimentoBomba.Periodo, MovimentoBomba.TipoCombustivel, MovimentoBomba.CodigoBomba, MovimentoBomba.SubCaixa)
                            
                        MovimentoBomba.TotalDesconto = lTotalDescontoNivelPreco
                        MovimentoBomba.TotalAcrescimo = lTotalAcrescimoNivelPreco
                        
                        If MovimentoBomba.Incluir Then
                            MovimentoBombaEscritorio.Empresa = g_empresa
                            MovimentoBombaEscritorio.Data = g_cfg_data_i 'lData
                            MovimentoBombaEscritorio.Periodo = g_cfg_periodo_i 'lPeriodo
                            MovimentoBombaEscritorio.SubCaixa = 999
                            MovimentoBombaEscritorio.CodigoBomba = Bomba.Codigo
                            MovimentoBombaEscritorio.Abertura = MovimentoBombaEscritorio.EncerranteBicoAnatesDataPeriodo(g_empresa, lData, i, 9, 999)
                            MovimentoBombaEscritorio.Encerrante = EncerranteAtual.Encerrante
                            MovimentoBombaEscritorio.QuantidadeSaida = MovimentoBomba.Encerrante - MovimentoBomba.Abertura
                            MovimentoBombaEscritorio.PrecoCusto = Bomba.PrecoCusto
                            MovimentoBombaEscritorio.PrecoVenda = Bomba.PrecoVenda
                            MovimentoBombaEscritorio.TipoCombustivel = Bomba.TipoCombustivel
                            MovimentoBombaEscritorio.NumeroTanque = Bomba.NumeroTanque
                            MovimentoBombaEscritorio.NumeroIlha = lIlha
                            MovimentoBombaEscritorio.TotalVendaBruto = MovimentoBombaEscritorio.QuantidadeSaida * MovimentoBombaEscritorio.PrecoVenda
                            MovimentoBombaEscritorio.TotalDesconto = lTotalDescontoNivelPreco
                            MovimentoBombaEscritorio.TotalAcrescimo = lTotalAcrescimoNivelPreco
                            
                            If Not MovimentoBombaEscritorio.Incluir Then
                                Call GravaAuditoria(1, Me.name, 28, "Não foi possível incluir o movimento de bomba-esc. do bico=" & i)
                                MsgBox "Não foi possível incluir o movimento de bomba-esc. do bico=" & i, vbInformation, "Duplicidade de Registro"
                            End If
                        Else
                            Call GravaAuditoria(1, Me.name, 28, "Não foi possível incluir o movimento de bomba do bico=" & i)
                            MsgBox "Não foi possível incluir o movimento de bomba do bico=" & i, vbInformation, "Duplicidade de Registro"
                        End If
                    End If
                End If
            Next
            If lData < g_cfg_data_i Then
                lData = g_cfg_data_i
            End If
            LoopIncluiMovBombaCaixa
            Call TelaAguarde("", False)
            NovoCupom
            TimerAutomacao.Enabled = True
            If txt_cliente.Enabled Then
                txt_cliente.SetFocus
            Else
                txt_funcionario_ponto.SetFocus
            End If
        End If
    End If
End Sub
'Private Sub mnuLeituraX_Click()
'    If l_flag_cupom_fiscal = "A" Then
'        MsgBox "Esta função não poderá ser executada com cupom aberto." & vbCrLf & "Feche o cupom que encontra-se aberto.", vbInformation + vbOKOnly, "Cupom Fiscal Aberto!"
'        Exit Sub
'    End If
'    Call GravaAuditoria(1, Me.name, 23, mnuLeituraX.Caption & " Func.:" & l_nome_funcionario)
'    'If (MsgBox("Imprime total por departamento?", vbQuestion + vbYesNo + vbDefaultButton2, "Totalizador por Departamento")) = vbYes Then
'    '    BemaRetorno = Bematech_FI_ImprimeDepartamentos()
'    'Else
'    '    Call CriaLogCupom("Bematech_FI_LeituraX")
'    '    BemaRetorno = Bematech_FI_LeituraX()
'    '    Call CriaLogCupom("Bematech_FI_LeituraX - BemaRetorno=" & BemaRetorno)
'    '    'ImprimeLeituraXCombustivel
'    'End If
'    If lImpBematech Then
'        Call CriaLogCupom("Bematech_FI_LeituraX")
'        BemaRetorno = Bematech_FI_LeituraX()
'        Call CriaLogCupom("Bematech_FI_LeituraX - BemaRetorno=" & BemaRetorno)
'    ElseIf lImpDaruma Then
'        BemaRetorno = Daruma_FI_LeituraX()
'    '01/04/2016
'    ElseIf lImpQuick Then
'        EcfQuickLeituraX
'    End If
'End Sub
Private Sub mnuMudaProximoTurno_Click()
    Dim xTipoVenda As String
    
    If (MsgBox("Deseja realmente mudar o período do caixa atual para o próximo periodo de caixa?", vbDefaultButton2 + vbYesNo + vbQuestion, "Fechamento de Caixa")) = vbYes Then
        Call GravaAuditoria(1, Me.name, 26, "Foi confirmado a mudança de período:" & lPeriodo)
        xTipoVenda = ReadINI("CUPOM FISCAL", "Tipo de Venda", gArquivoIni)
        If xTipoVenda = "AUTOMACAO/CONVENIENCIA" Then
            If LiberacaoDigitacao.LocalizarCodigo(g_empresa, 2) Then
                If LiberacaoDigitacao.PeriodoInicial = gQtdPeriodo Then
                    LiberacaoDigitacao.DataInicial = LiberacaoDigitacao.DataInicial + 1
                    LiberacaoDigitacao.DataFinal = LiberacaoDigitacao.DataFinal + 1
                    LiberacaoDigitacao.PeriodoInicial = 0
                    LiberacaoDigitacao.PeriodoFinal = 0
                End If
                LiberacaoDigitacao.PeriodoInicial = LiberacaoDigitacao.PeriodoInicial + 1
                LiberacaoDigitacao.PeriodoFinal = LiberacaoDigitacao.PeriodoFinal + 1
                If LiberacaoDigitacao.Alterar(g_empresa, 2) Then
                    If LiberacaoDigitacao.LocalizarCodigo(g_empresa, 3) Then
                        If LiberacaoDigitacao.PeriodoInicial = gQtdPeriodo Then
                            LiberacaoDigitacao.DataInicial = LiberacaoDigitacao.DataInicial + 1
                            LiberacaoDigitacao.DataFinal = LiberacaoDigitacao.DataFinal + 1
                            LiberacaoDigitacao.PeriodoInicial = 0
                            LiberacaoDigitacao.PeriodoFinal = 0
                        End If
                        LiberacaoDigitacao.PeriodoInicial = LiberacaoDigitacao.PeriodoInicial + 1
                        LiberacaoDigitacao.PeriodoFinal = LiberacaoDigitacao.PeriodoFinal + 1
                        If Not LiberacaoDigitacao.Alterar(g_empresa, 3) Then
                            MsgBox "Erro ao alterar a Liberação de Digitação!", vbInformation, "Registro não Alterado!"
                        End If
                    End If
                Else
                    MsgBox "Erro ao alterar a Liberação de Digitação!", vbInformation, "Registro não Alterado!"
                End If
            End If
        Else
            If LiberacaoDigitacao.LocalizarCodigo(g_empresa, 2) Then
                If LiberacaoDigitacao.PeriodoInicial = gQtdPeriodo Then
                    LiberacaoDigitacao.DataInicial = LiberacaoDigitacao.DataInicial + 1
                    LiberacaoDigitacao.DataFinal = LiberacaoDigitacao.DataFinal + 1
                    LiberacaoDigitacao.PeriodoInicial = 0
                    LiberacaoDigitacao.PeriodoFinal = 0
                End If
                LiberacaoDigitacao.PeriodoInicial = LiberacaoDigitacao.PeriodoInicial + 1
                LiberacaoDigitacao.PeriodoFinal = LiberacaoDigitacao.PeriodoFinal + 1
                If Not LiberacaoDigitacao.Alterar(g_empresa, 2) Then
                    MsgBox "Erro ao alterar a Liberação de Digitação!", vbInformation, "Registro não Alterado!"
                End If
            End If
        End If
        g_cfg_periodo_i = LiberacaoDigitacao.PeriodoInicial
        g_cfg_periodo_f = LiberacaoDigitacao.PeriodoFinal
        g_cfg_data_i = LiberacaoDigitacao.DataInicial
        g_cfg_data_f = LiberacaoDigitacao.DataFinal
        Me.Caption = "Nota Fiscal do Consumidor Eletrônica - " & l_nome_funcionario & " | Caixa: " & Val(g_cfg_periodo_i) & " Em: " & Format(g_cfg_data_i, "dd/mm/yyyy")
    Else
        Call GravaAuditoria(1, Me.name, 26, "Não foi confirmado o fechamento de caixa:" & lPeriodo)
    End If
End Sub
Private Sub mnuNFCeCupomComplementar_Click()
    If g_automacao = True Then
        Call CriaLogCupom("NFCe Auto: A Emissão da NFCe Complementar Automação Foi Chamada.")
        EmissaoNFCeComplementarAuto.Show 1
    Else
        Call CriaLogCupom("NFCe Auto: A Emissão da NFCe Complementar Foi Chamada.")
        EmissaoNFCeComplementar.Show 1
    End If
    AtualizaBombasAbastecimento
    AutomacaoMostraBicos
End Sub
Private Sub mnuPontoFuncionario_Click()
    'Abre o cupom fiscal
'    Call GravaAuditoria(1, Me.name, 23, mnuPontoFuncionario.Caption & " Func.:" & l_nome_funcionario)
'    Call CriaLogCupom("Bematech_FI_AbreCupom")
'    BemaRetorno = Bematech_FI_AbreCupom("")
'    Call CriaLogCupom("Bematech_FI_AbreCupom - BemaRetorno=" & BemaRetorno)
'    'Imprime Produto
'    Call CriaLogCupom("Bematech_FI_VendeItemDepartamento(... l_nome_funcionario ...) - l_nome_funcionario=" & l_nome_funcionario)
'    BemaRetorno = Bematech_FI_VendeItemDepartamento(Format(l_codigo_funcionario, "#,##0"), l_nome_funcionario, "II", "000000010", "0001000", "0000000000", "0000000000", "05", "PO")
'    Call CriaLogCupom("Bematech_FI_VendeItemDepartamento - BemaRetorno=" & BemaRetorno)
'    'Cancela o cupom fiscal
'    Call CriaLogCupom("Bematech_FI_CancelaCupom")
'    BemaRetorno = Bematech_FI_CancelaCupom
'    Call CriaLogCupom("Bematech_FI_CancelaCupom - BemaRetorno=" & BemaRetorno)
'    NovoCupom
End Sub


'Private Sub mnuReducaoZ_Click()
'    If l_flag_cupom_fiscal = "A" Then
'        MsgBox "Esta função não poderá ser executada com cupom aberto." & vbCrLf & "Feche o cupom que encontra-se aberto.", vbInformation + vbOKOnly, "Cupom Fiscal Aberto!"
'        Exit Sub
'    End If
'    If (MsgBox("Deseja realmente imprimir a redução Z?", vbQuestion + vbYesNo + vbDefaultButton2, "Impressão de Redução Z!")) = vbNo Then
'        Exit Sub
'    End If
'    Call GravaAuditoria(1, Me.name, 23, mnuReducaoZ.Caption & " Func.:" & l_nome_funcionario)
'    'Call ImprimeReducaoZ
'End Sub
Private Sub mnuSenha_Click()
    Call GravaAuditoria(1, Me.name, 23, mnuSenha.Caption & " Func.:" & l_nome_funcionario)
    MostraTelaSenha
    
'    Call AbilitaMenu(False)
'    frm_ponto.Enabled = True
'    frm_ponto.ZOrder 0
'    txt_funcionario_ponto = ""
'    dtcboFuncionario.BoundText = 0
'    txt_senha_ponto.Text = ""
'    'mnuLeituraX.Enabled = False
'    'cmd_senha.Enabled = False
'    mnuPontoFuncionario.Enabled = False
'    frmDados.Enabled = False
'    frm_fechamento_cupom.Enabled = False
'    txt_cupom_fiscal.Enabled = False
'    txt_funcionario_ponto.SetFocus
'    If lDescontoPostoAki Then
'        Call PreencheCboFormaPagamento
'    End If
End Sub
Private Sub mnuTCS_Click()
    Call GravaAuditoria(1, Me.name, 24, "Funcionário:" & l_nome_funcionario)
    Call AtivaDesativaTimer(False)
    AtualizaPrecoTCS
    Call AtivaDesativaTimer(True)
    mnuSenha_Click
End Sub
Private Sub mnuVisualizaVenda_Click()
    BuscaPeriodo
    Call GravaAuditoria(1, Me.name, 23, mnuVisualizaVenda.Caption & " Func.:" & l_nome_funcionario)
    If g_automacao Then
        If (MsgBox("Deseja realmente visualizar venda do dia?", vbQuestion + vbYesNo + vbDefaultButton2, "Visualiza Venda")) = vbYes Then
            TimerAutomacao.Enabled = False
            If (MsgBox("Acumular vendas do dia?", vbQuestion + vbYesNo + vbDefaultButton2, "Vendas Atual")) = vbYes Then
                g_string = "ConsultaVenda|@|" & lData & "|@|" & 0 & "|@|"
            Else
                g_string = "ConsultaVenda|@|" & lData & "|@|" & lPeriodo & "|@|"
            End If
            g_string = ""
            NovoCupom
            TimerAutomacao.Enabled = True
            If txt_cliente.Enabled Then
                txt_cliente.SetFocus
            Else
                txt_funcionario_ponto.SetFocus
            End If
        End If
    End If
End Sub
Private Sub MSFlexGrid_Click()
    MarcaCelulaFlexGrid
End Sub
Private Sub MSFlexGrid_DblClick()
    MarcaCelulaFlexGrid
End Sub
Private Sub MSFlexGrid_SelChange()
    MarcaCelulaFlexGrid
End Sub
Private Sub Timer2_Timer()
    Dim x_mensagem As String
    
    'Sai do Horário de Verão
'    If UCase(g_nome_empresa) Like "*BOSQUE*" Then
'        If l_flag_cupom_fiscal = "F" Then
'            If Date = CDate("21/03/2006") Then
'                If Time >= "22:45:00" Then
'                    SaiHorarioVerao
'                    Exit Sub
'                End If
'            End If
'        End If
'    End If
    x_mensagem = l_mensagem
    x_tempo = x_tempo + 1
    If x_tempo = 1 Then
        lbl_mensagem = x_mensagem
    Else
        If x_tempo <= Len(x_mensagem) Then
            lbl_mensagem = Space(1) & Mid(x_mensagem, x_tempo, Len(x_mensagem) - x_tempo)
        Else
            x_tempo = 0
        End If
    End If
    
    If Len(gStringChamadaSangria) > 0 And l_flag_cupom_fiscal = "F" Then
        'Call TestaConexao(40, "Timer2 - gStringChamadaSangria=" & gStringChamadaSangria)
        If ConfiguracaoDiversa.LocalizarCodigo(1, "Imprime Sangria") Then
            'Call TestaConexao(41, "Timer2")
            If ConfiguracaoDiversa.Verdadeiro = True Then
                'Call TestaConexao(42, "Timer2")
                Dim xValor As String
                Dim xPeriodo As String
                xValor = RetiraString(2, gStringChamadaSangria)
                xPeriodo = RetiraString(3, gStringChamadaSangria)
                gStringChamadaSangria = ""
                Call GeraSangriaTermica(xValor, xPeriodo)
            End If
        End If
        'Call TestaConexao(43, "Timer2")
        gStringChamadaSangria = ""
    End If

End Sub

Private Sub GeraSangriaTermica(ByVal pValor As String, ByVal xPeriodo As String, Optional ByVal pQuantidadeVias As Integer = 2)
    Dim xString As String
    Dim xString2 As String
    Dim xLinha As String
    Dim xData As Date
    Dim i As Integer
    Dim ContadorImpressao As Integer
    
    
    
    Call CriaLogCupom("[GeraSangriaTermica] Valor = " & pValor & " Período = " & xPeriodo & " Qtd. Vias = " & pQuantidadeVias)
    
    '                  1         2         3         4       4
    '          123456789012345678901234567890123456789012345678
    'xLinha = "                                                "
    
    If pQuantidadeVias <= 0 Then
        Exit Sub
    End If
    
    DefineImpressoraTermicaComoPadrao
    
For ContadorImpressao = 1 To pQuantidadeVias
    
    ImpTermicaAbreRelatorio
    Call ImpTermicaCabecalhoPosto(0)
    
    
    xString2 = "VIA " & ContadorImpressao
    xString = xString + xString2
    Call ImpTermicaImprimeDados(xString2, True)
    
    'PULO DE 1 LINHA
    xString2 = "                                                "
    xString = xString + xString2
    Call ImpTermicaImprimeDados(xString2, True)
    
    xString2 = "EMISSÃO DE SANGRIA"
    xString = xString + xString2
    Call ImpTermicaImprimeDados(xString2, True)
    
    xData = lData
    
    xString2 = "Data........:                                  "
    Mid(xString2, 15, 34) = CStr(xData) '"Data.........: " + CStr(xData)
    xString = xString + xString2
    Call ImpTermicaImprimeDados(xString2, True)
    
    xString2 = "Funcionario.:                                  "
    Mid(xString2, 15, 34) = CStr(l_nome_funcionario)
    xString = xString + xString2
    Call ImpTermicaImprimeDados(xString2, True)

    xString2 = "Valor.......:                                  "
    Mid(xString2, 15, 34) = CStr(pValor)
    xString = xString + xString2
    Call ImpTermicaImprimeDados(xString2, True)
    
    xString2 = "Periodo.....:                                  "
    Mid(xString2, 15, 34) = CStr(xPeriodo)
    xString = xString + xString2
    Call ImpTermicaImprimeDados(xString2, True)
    
    'PULO DE 2 LINHAS
    xString2 = "                                                "
    xString = xString + xString2
    Call ImpTermicaImprimeDados(xString2, True)
    Call ImpTermicaImprimeDados(xString2, True)
    
        
    Mid(xString2, 1, 48) = "ASSINATURA: ____________________________________"
    xString = xString + xString2
    Call ImpTermicaImprimeDados(xString2, True)

    
    xString2 = "                                                "
    i = Len(l_nome_funcionario)
    'Mid(xString2, ((48 - i) / 2), i) = l_nome_funcionario
    Mid(xString2, ((36 - i) / 2), i) = l_nome_funcionario
    xString = xString + xString2
    Call ImpTermicaImprimeDados(xString2, True)
    
    'PULO DE 2 LINHAS
    xString2 = "                                                "
    xString = xString + xString2
    Call ImpTermicaImprimeDados(xString2, True)
    Call ImpTermicaImprimeDados(xString2, True)
    
    ImpTermicaFechaRelatorio ("Emissão de Sangria")

Next
    
    Call CriaLogCupom(Time & " - Emissão de Sangria: xString=" & xString)
    
    
    Exit Sub

End Sub


Private Sub TimerAguarde_Timer()
    
    lContadorAguarde = lContadorAguarde - 1
    lblContadorAguarde.Caption = lContadorAguarde
    DoEvents
    If lContadorAguarde <= 0 Then
        Call AtivaDesativaAguarde("", False)
    End If
End Sub

Private Sub TimerAutomacao_Timer()
    'Atualiza Hora Placa             - lAutomacaoFlag = 0 e 11
    'Pede Abastecimento Efetuado     - lAutomacaoFlag = 6 e 1
    'Pede Abastecimento em Andamento - lAutomacaoFlag = 2,3,4 e 5
    Dim x_string As String
    Dim xPassoDebugarErro As String
    
    On Error GoTo FileError
    
    xPassoDebugarErro = "1"
    If lAutomacaoFlag = 0 Then
        xPassoDebugarErro = "2"
        'AutomacaoAtualizaRelogioPlaca
        lAutomacaoFlag = 11
        'Agora que nao le automacao muda pra 1
        lAutomacaoFlag = 1
        Exit Sub
    ElseIf lAutomacaoFlag = 1 Then
        
        If lUfEmpresa = "PA" Then
            AtualizaBombasAbastecimento '-ALEX_HOJE
        End If
        
        xPassoDebugarErro = "3"
'        If l_flag_cupom_fiscal = "F" Then
'            If lExisteMudancaHorarioVerao Then
'                If Date >= MovHorarioVerao.DataParaInicioBloqueio Then
'                    If Format(Time, "HH:mm:ss") >= Format(MovHorarioVerao.HoraParaInicioBloqueio, "HH:mm:ss") Then
'                        MudaHorarioVeraoAutomatico
'                    End If
'                End If
'            End If
'        End If
    ElseIf lAutomacaoFlag > 1 And lAutomacaoFlag < 6 Then
        xPassoDebugarErro = "4"
        'AutomacaoPedeAbastecimentoEmAndamento
    End If
    xPassoDebugarErro = "5"
    If lAutomacaoFlag > 0 And lAutomacaoFlag < 6 Then
        xPassoDebugarErro = "6"
        lAutomacaoFlag = lAutomacaoFlag + 1
    ElseIf lAutomacaoFlag = 11 Then
        xPassoDebugarErro = "7"
        lAutomacaoFlag = 1
        Exit Sub
    End If
    xPassoDebugarErro = "10"
    If lAutomacaoFlag = 6 Then
        xPassoDebugarErro = "11"
        lAutomacaoFlag = 1
    Else
        xPassoDebugarErro = "12"
    End If
    xPassoDebugarErro = "13"
    
    If lUfEmpresa = "PA" Then
        AutomacaoMostraBicos  '- ALEX_HOJE
    End If
    
    xPassoDebugarErro = "14"
    Exit Sub

FileError:
    Call CriaLogCupom("Cupom Fiscal: ERRO TimerAutomacao: " & Error & " - xPassoDebugarErro:" & xPassoDebugarErro & " - Err: " & Err)
    Exit Sub
End Sub

Private Sub TimerIdentFid_Timer()
    Dim i As Integer
    Dim xString As String
    
    xString = MSCommIdentFid.Input
    If Len(xString) = 23 Then
        TimerIdentFid.Enabled = False
        If MSCommIdentFid.PortOpen = True Then
            MSCommIdentFid.PortOpen = False
        End If
        If CartaoAbastecimento.LocalizarNumeroCartao(g_empresa, Mid(xString, 4, 16)) Then
            txt_funcionario_ponto.Text = CartaoAbastecimento.CodigoFuncionario
            dtcboFuncionario.BoundText = CartaoAbastecimento.CodigoFuncionario
            If Funcionario.LocalizarCodigo(g_empresa, CartaoAbastecimento.CodigoFuncionario) Then
                If Usuario.LocalizarCodigo(Funcionario.CodigoUsuario) Then
                    txt_senha_ponto.Text = DesKriptografa(Funcionario.Senha)
                    l_senha_funcionario = Funcionario.Senha
                    cmd_ok_ponto_Click
                Else
                    txt_senha_ponto.SetFocus
                End If
            Else
                txt_senha_ponto.SetFocus
            End If
        End If
    ElseIf Len(xString) > 0 Then
        MsgBox "xstring=" & xString & vbCrLf & "Cartao=" & Mid(xString, 4, 16) & vbCrLf & "Tamanho=" & Len(xString)
    End If
End Sub

Private Sub txt_cliente_GotFocus()
    txt_cliente.BackColor = &H800000
    txt_cliente.ForeColor = &H80000005
    l_mensagem = Space(165) & "Digite 0(zero) para cliente não cadastrado.  |  Informe o código do cliente.  |  Tecle enter para informar o nome do cliente.  |  Tecle F12 para cancelar o último cupom fiscal."
    lOrigemFocus = "txt_cliente"
    If Val(txt_cliente.Text) = 0 Then
        txt_cliente.Text = "0"
    End If
    txt_cliente.SelStart = 0
    txt_cliente.SelLength = Len(txt_cliente.Text)
End Sub
Private Sub txt_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Dim xContinua As Boolean
    
    'F12
    If KeyCode = 123 Then
        KeyCode = 0
        MsgBox "Função pra NFCe não implementada!", vbOKOnly + vbInformation, "Operação Não Aceita."
        Exit Sub
        xContinua = True
        If lImpBematech Then
            'Call CriaLogCupom("Bematech_FI_FlagsFiscais(i)")
            'BemaRetorno = Bematech_FI_FlagsFiscais(i)
            'Call CriaLogCupom("Bematech_FI_FlagsFiscais(i)=" & i & " - BemaRetorno=" & BemaRetorno)
            'If i = 32 Or i = 36 Then
                xContinua = True
            'Else
            '    Call GravaAuditoria(1, Me.name, 25, "Tentativa não permitida por tempo excedido. ECF:" & lNumeroCupom)
            '    MsgBox "Cancelamento do último cupom não permitido(3)." & Chr(10) & "Tempo excedido!", vbInformation, "Cancelamento Não Permitido!"
            '    xContinua = False
            'End If
            'aquiaquiaqui 01/04
        ElseIf lImpDaruma Then
            xContinua = True
        End If
        If xContinua Then
            If (MsgBox("Deseja cancelar o último cupom fiscal?", vbYesNo + vbDefaultButton2 + vbQuestion, "Cancela Cupom Fiscal")) = vbYes Then
                If lImpBematech Then
'                        Call CriaLogCupom("Bematech_FI_FlagsFiscais(i)")
'                        BemaRetorno = Bematech_FI_FlagsFiscais(i)
'                        Call CriaLogCupom("Bematech_FI_FlagsFiscais(i)=" & i & " - BemaRetorno=" & BemaRetorno)
'                        If i = 32 Or i = 36 Then
'                            xContinua = True
'                        Else
'                            Call GravaAuditoria(1, Me.name, 25, "Não permitido por tempo excedido. ECF:" & lNumeroCupom)
'                            MsgBox "Cancelamento do último cupom não permitido(4)." & Chr(10) & "Tempo excedido!", vbInformation, "Cancelamento Não Permitido!"
'                            xContinua = False
'                        End If
                ElseIf lImpDaruma Then
                    xContinua = True
                End If
            Else
                xContinua = False
            End If
        End If
        If xContinua Then
            Call GravaAuditoria(1, Me.name, 25, "Inicio. ECF:" & lNumeroCupom)
            If CancelamentoCupomFiscal Then
                NovoCupom
                Call MontaCupomVideo(lNumeroCupom, lData)
                mnuSenha_Click
            End If
        End If
    End If
End Sub
Private Sub txt_cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dtcboCliente.SetFocus
    ElseIf KeyAscii = 27 Then
        KeyAscii = 0
        Unload Me
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_cliente_LostFocus()
    txt_cliente.BackColor = &H80000005
    txt_cliente.ForeColor = &H80000008
    l_codigo_cliente = Val(txt_cliente.Text)
    If Val(txt_cliente.Text) > 0 Then
        If Cliente.LocalizarCodigo(CLng(txt_cliente.Text)) Then
            If Cliente.Inativo = True Then
                MsgBox "O cliente " & Trim(Cliente.RazaoSocial) & " está inativo.", vbInformation, "Cliente Inativo!"
                txt_cliente.SetFocus
                Exit Sub
            Else
                dtcboCliente.BoundText = Val(txt_cliente.Text)
                dtcboCliente_LostFocus
                Exit Sub
            End If
        Else
            MsgBox "Cliente não cadastro.", vbInformation, "Atenção!"
            dtcboCliente.BoundText = ""
            txt_cliente.SetFocus
        End If
    ElseIf txt_cliente.Text = "0" Then
        dtcboCliente.BoundText = 0
        'dtcboCliente_LostFocus
    End If
End Sub
Private Sub txt_cpf_GotFocus()
    If Len(Trim(txt_cpf.Text)) > 0 Then
        txt_cpf.Text = fDesmascaraNumeroString(txt_cpf.Text)
    End If
    txt_cpf.SelStart = 0
    txt_cpf.SelLength = Len(txt_cpf.Text)
End Sub
Private Sub txt_cpf_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_nome_cliente.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_cpf_LostFocus()
    Dim xCpfDesmascarado As String
    
    xCpfDesmascarado = txt_cpf.Text
    If Len(txt_cpf.Text) = 11 Then
        txt_cpf.Text = fMascaraCPF(txt_cpf.Text)
    ElseIf Len(txt_cpf.Text) = 14 Then
        txt_cpf.Text = fMascaraCNPJ(txt_cpf.Text)
    End If
    If cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 2 Or cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 3 Then
        If txt_cpf.Text <> "" Then
            g_string = MovimentoCheque.LocalizarCpfCnpj(xCpfDesmascarado)
            If g_string <> "" Then
                txt_nome_cliente.Text = RetiraGString(1)
                txt_telefone.Text = fMascaraTelefone(RetiraGString(2))
                'txtBanco.Text = RetiraGString(3)
                'txtAgencia.Text = RetiraGString(4)
                'txtCpfCnpj.Text = RetiraGString(5)
                'txt_conta.Text = RetiraGString(6)
                txt_numero_cheque.SetFocus
                g_string = ""
            Else
                txt_nome_cliente.Text = ""
                txt_telefone.Text = ""
            End If
        End If
    End If
End Sub
Private Sub txt_funcionario_ponto_GotFocus()
    txt_funcionario_ponto.BackColor = &H800000
    txt_funcionario_ponto.ForeColor = &H80000005
    l_mensagem = Space(165) & "Informe o código do funcionário."
    txt_funcionario_ponto.SelStart = 0
    txt_funcionario_ponto.SelLength = Len(txt_funcionario_ponto)
    Me.Caption = "Nota Fiscal do Consumidor Eletrônica (NFCe)"
    If lCaixaIndividual And lPortaRfid > 0 Then
        If MSCommIdentFid.PortOpen = False Then
            MSCommIdentFid.CommPort = lPortaRfid
            MSCommIdentFid.Settings = "9600,n,8,1"
            MSCommIdentFid.PortOpen = True
            MSCommIdentFid.InBufferCount = 0
            TimerIdentFid.Enabled = True
            TimerIdentFid.Interval = 500
        End If
    End If
End Sub
Private Sub txt_funcionario_ponto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dtcboFuncionario.SetFocus
    ElseIf KeyAscii = 27 Then
        KeyAscii = 0
        Unload Me
    'Crtl + V
    ElseIf KeyAscii = 22 Then
        ZZTotalizaCupomAbertoNoBanco
    'Crtl + C
    ElseIf KeyAscii = 3 And lImpQuick Then
        KeyAscii = 0
        'GeraCat52DataRegis
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_funcionario_ponto_LostFocus()
    If MSCommIdentFid.PortOpen = True Then
        MSCommIdentFid.PortOpen = False
        TimerIdentFid.Enabled = False
    End If
    txt_funcionario_ponto.BackColor = &H80000005
    txt_funcionario_ponto.ForeColor = &H80000008
    If Val(txt_funcionario_ponto) > 0 Then
        If Funcionario.LocalizarCodigo(g_empresa, Val(txt_funcionario_ponto.Text)) Then
            If Funcionario.Situacao = "I" Then
                MsgBox "O funcionário " & Trim(Funcionario.Nome) & " está inativo.", vbInformation, "Atenção!"
                txt_funcionario_ponto.SetFocus
                Exit Sub
            Else
                dtcboFuncionario.BoundText = Funcionario.Codigo
                l_senha_funcionario = Funcionario.Senha
                If Usuario.LocalizarCodigo(Funcionario.CodigoUsuario) Then
                Else
                    MsgBox "Funcionário sem código do usuário no cadastrao.", vbInformation, "Erro "
                    cmd_cancelar_ponto_Click
                    Exit Sub
                End If
                If txt_senha_ponto.Enabled Then
                    txt_senha_ponto.SetFocus
                End If
            End If
        Else
            MsgBox "Funcionário não cadastrado.", vbInformation, "Atenção!"
            txt_funcionario_ponto.SetFocus
            Exit Sub
        End If
    End If
End Sub
Private Sub txt_kilometragem_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_numero_nota_abastecimento.SetFocus
    End If
End Sub
Private Sub txt_nome_cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt_valor_desconto.Enabled = True Then
            txt_valor_desconto.SetFocus
        Else
            txt_valor_recebido.SetFocus
        End If
    End If
End Sub
Private Sub txt_numero_cheque_GotFocus()
    l_mensagem = Space(165) & "Informe o número do cheque recebido."
End Sub
Private Sub txt_numero_cheque_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_telefone.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_numero_nota_abastecimento_GotFocus()
    txt_numero_nota_abastecimento.SelStart = 0
    txt_numero_nota_abastecimento.SelLength = Len(txt_numero_nota_abastecimento.Text)
End Sub
Private Sub txt_numero_nota_abastecimento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cmdIniciaProcessoNFCe.Enabled Then
            cmdIniciaProcessoNFCe.SetFocus
        End If
    End If
End Sub
Private Sub txt_observacao_2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt_valor_desconto.Enabled = True Then
            txt_valor_desconto.SetFocus
        Else
            txt_valor_recebido.SetFocus
        End If
    End If
End Sub
Private Sub txt_observacao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_observacao_2.SetFocus
    End If
End Sub
Private Sub txt_placa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_kilometragem.SetFocus
    End If
End Sub
Private Sub txt_produto_GotFocus()
    If Val(txt_cliente.Text) > 0 Then
        
        If lOrigemFocus = "dtcboCliente" Or lOrigemFocus = "dtcboCliente" Or lOrigemFocus = "txt_cliente" Or lOrigemFocus = "txt_cliente_conveniado" Then 'NEW 07/04
            lCodigoVeiculo = 0 'NEW 07/04
            If dtcboCliente.BoundText <> "" Then 'NEW 07/04
                SelecionaVeiculoCliente (Cliente.Codigo) 'NEW 07/04
            End If 'NEW 07/04
        End If 'NEW 07/04
    
        If lOrigemFocus = "dtcboCliente" Or lOrigemFocus = "txt_cliente" Or lOrigemFocus = "txt_cliente_conveniado" Then
            VerificaClienteEmAtraso
        End If
    End If
    lOrigemFocus = "txt_produto"
    txt_produto.BackColor = &H800000
    txt_produto.ForeColor = &H80000005
    'l_mensagem = Space(165) & "Informe o código do produto.  |  Tecle enter para informar o nome do produto.  |  Tecle F8 para cancelar ítem à escolher.  |  Tecle F10 para informar a forma de pagamento.  |  Tecle F12 para cancelar o último cupom fiscal. "
    l_mensagem = Space(165) & "Informe o código do produto.  |  Tecle enter para informar o nome do produto.  |  Tecle F9 para Abortar a venda.  |  Tecle F10 para informar a forma de pagamento. "
    txt_produto.SelStart = 0
    txt_produto.SelLength = Len(txt_produto)
    lInformaFormaPagamento = False
    lCodigoBarra = False
End Sub
Private Sub txt_produto_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Dim xContinua As Boolean
    
    'F8 Cancela ítem à escolher do cupom fiscal
    
    If KeyCode = 119 Then
        KeyCode = 0
        MsgBox "Função pra NFCe não implementada!", vbOKOnly + vbInformation, "Operação Não Aceita."
        Exit Sub

'        If lOrdem > 1 Then
'            If (MsgBox("Deseja cancelar o último item?", vbYesNo + vbDefaultButton2 + vbQuestion, "Cancela Item")) = vbYes Then
'                Call GravaAuditoria(1, Me.name, 25, "Inicio. ECF:" & lNumeroCupom & " Ordem:" & lOrdem - 1)
'                CancelamentoCupomFiscalItem
'                NovoCupom
'                Call MontaCupomVideo(lNumeroCupom, lData)
'            End If
'        Else
'            MsgBox "Não existe ítem a ser cancelado!", vbInformation, "Operação não aceita."
'        End If
    End If
    
    'F9 - abortar NFCe
    If KeyCode = 120 Then
        Call AbortarNFCeAtual
        Exit Sub
    End If
    'F10 Fecha Cupom Fiscal
    If KeyCode = 121 Then
        KeyCode = 0
        If l_flag_cupom_fiscal = "A" Then
            lInformaFormaPagamento = True
            Call CriaLogMonitoraBug("txt_produto KeyDown F10", "Movimento_NfceAuto", "Chamou CancelaCupom")
            CancelaCupom
        End If
    End If
    'F12
    If KeyCode = 123 Then
        KeyCode = 0
        MsgBox "Função pra NFCe não implementada!", vbOKOnly + vbInformation, "Operação Não Aceita."
        Exit Sub
        xContinua = True
'        If lExisteImpressora Then
'            If lImpBematech Then
'                Call CriaLogCupom("Bematech_FI_FlagsFiscais(i)")
'                BemaRetorno = Bematech_FI_FlagsFiscais(i)
'                Call CriaLogCupom("Bematech_FI_FlagsFiscais(i) - i=" & i & " - BemaRetorno=" & BemaRetorno)
'                If i = 32 Or i = 36 Then
'                    xContinua = True
'                Else
'                    Call GravaAuditoria(1, Me.name, 25, "Tentativa não permitida por tempo excedido. ECF:" & lNumeroCupom)
'                    MsgBox "Cancelamento do último cupom não permitido(5)." & Chr(10) & "Tempo excedido!", vbInformation, "Cancelamento Não Permitido!"
'                    xContinua = False
'                End If
'            ElseIf lImpDaruma Then
'                xContinua = True
'            End If
'        End If
        If xContinua Then
            xContinua = False
            If (MsgBox("Deseja cancelar o último cupom fiscal?", vbYesNo + vbDefaultButton2 + vbQuestion, "Cancela Cupom Fiscal")) = vbYes Then
'                If lExisteImpressora Then
'                    If lImpBematech Then
'                        Call CriaLogCupom("Bematech_FI_FlagsFiscais(i)")
'                        BemaRetorno = Bematech_FI_FlagsFiscais(i)
'                        Call CriaLogCupom("Bematech_FI_FlagsFiscais(i) - i=" & i & " - BemaRetorno=" & BemaRetorno)
'                        If i = 32 Or i = 36 Then
'                            xContinua = True
'                        Else
'                            Call GravaAuditoria(1, Me.name, 25, "Não permitido por tempo excedido. ECF:" & lNumeroCupom)
'                            MsgBox "Cancelamento do último cupom não permitido(6)." & Chr(10) & "Tempo excedido!", vbInformation, "Cancelamento Não Permitido!"
'                            xContinua = False
'                        End If
'                    ElseIf lImpDaruma Then
'                        xContinua = True
'                    End If
'                End If
            End If
        End If
        If xContinua Then
            Call GravaAuditoria(1, Me.name, 25, "Inicio. ECF:" & lNumeroCupom)
            If CancelamentoCupomFiscal Then
                NovoCupom
                Call MontaCupomVideo(lNumeroCupom, lData)
                mnuSenha_Click
            End If
        End If
    End If
End Sub
Private Sub txt_produto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dtcboProduto.SetFocus
    'Crtl + G Abre a Gaveta
    ElseIf KeyAscii = 7 Then
        Dim xRetorno As Integer
        KeyAscii = 0
        Call CriaLogCupom("Bematech_FI_AcionaGaveta")
        xRetorno = Bematech_FI_AcionaGaveta
        Call CriaLogCupom("Bematech_FI_AcionaGaveta - xRetorno=" & xRetorno)
    End If
    Call ValidaInteiroQtd(KeyAscii)
End Sub
Private Sub txt_produto_LostFocus()
    Dim i As Integer
    Dim xValorTotal As Currency
    Dim xQuantidade As Integer
    
    txt_produto.BackColor = &H80000005
    txt_produto.ForeColor = &H80000008
    
    xQuantidade = 1
    If txt_produto.Text <> "" Then
        If lLoja Then
            txt_quantidade.Text = 1
            For i = 1 To Len(txt_produto.Text)
                If Mid(txt_produto.Text, i, 1) = "*" Then
                    xQuantidade = Mid(txt_produto.Text, 1, i - 1)
                    txt_produto.Text = Mid(txt_produto.Text, i + 1, Len(txt_produto.Text) - i)
                    txt_quantidade.Text = xQuantidade
                    Exit For
                End If
            Next
        End If
    End If
    
    
    lCodigoBarra = False
    xValorTotal = 0
    'Codigo de Barra de Balanca/Preco
    If Len(txt_produto.Text) > 10 Then
        If Mid(txt_produto.Text, 1, 1) = "2" Then
            lCodigoBarra = True
            xValorTotal = fValidaValor(Mid(txt_produto.Text, 6, 5) & "," & Mid(txt_produto.Text, 11, 2))
            txt_produto.Text = CLng(Mid(txt_produto.Text, 2, 4))
            If Produto.LocalizarCodigoBarra(txt_produto.Text) Then
                txt_produto.Text = Produto.Codigo
            Else
                MsgBox "Codigo de Barra não cadastrado!", vbInformation, "Erro de Leitura"
                txt_produto.Text = ""
                txt_produto.SetFocus
                Exit Sub
            End If
        End If
    End If
    
    'Codigo de Barra
    If Len(txt_produto.Text) > 5 Then
        lCodigoBarra = True
        If Produto.LocalizarCodigoBarra(txt_produto.Text) Then
            txt_produto.Text = Produto.Codigo
        Else
            MsgBox "Codigo de Barra não cadastrado!", vbInformation, "Erro de Leitura"
            txt_produto.Text = ""
            txt_produto.SetFocus
            Exit Sub
        End If
    End If
    
    If Val(txt_produto.Text) > 0 Then
        If Produto.LocalizarCodigo(CLng(txt_produto.Text)) Then
            If lExigeNCM = True Then
                If LocalizarNCM(0, Trim(Produto.CodigoNCM)) = False Then
                    txt_produto.SetFocus
                    Exit Sub
                End If
            End If
            
            'alex - NFCe
            'If Aliquota.LocalizarCodigo(lSerieECF, Produto.CodigoAliquota) Then
            'End If

            'alex - NFCe
            If Not Aliquota.LocalizarCodigo(lSerieECF, Produto.CodigoAliquota) Then
               MsgBox "Aliquota não cadastrada!", vbInformation, "Erro de Integridade!"
               txt_produto.SetFocus
               Exit Sub
            End If
            
'            LUIZ 28/11/2016 não é obrigatorio informar o cest
'            If (Aliquota.CodigoFiscal = CODIGO_FISCAL_SUBSTITUICAO) And Trim(Produto.CEST) = Empty Then
'                MsgBox "O produto " & Trim(Produto.Nome) & vbCrLf & "é Substituição Tributária e está cadastrado sem CEST.", vbInformation, "Produto Sem CEST!"
'                txt_produto.SetFocus
'                Exit Sub
'            End If


            
            If Produto.Inativo = True Then
                MsgBox "O produto " & Trim(Produto.Nome) & " está inativo.", vbInformation, "Produto Inativo!"
                txt_produto.SetFocus
                Exit Sub
            ElseIf Produto.ImprimeCupomFiscal = False Then
                MsgBox "O produto " & Trim(Produto.Nome) & " está configurado para não imprimir cupom fiscal.", vbInformation, "Impressão de Cupom Não Autorizada!"
                txt_produto.SetFocus
                Exit Sub
            Else
                dtcboProduto.BoundText = CLng(txt_produto.Text)
                If Estoque.LocalizarCodigo(g_empresa, CLng(txt_produto.Text)) Then
                    If Estoque.PrecoVenda <> 0 Then
                        txt_valor_unitario.Text = Format(Estoque.PrecoVenda, "###,##0.0000")
                    Else
                        txt_valor_unitario.Text = Format(Produto.PrecoVenda, "###,##0.0000")
                    End If
                Else
                    MsgBox "Estoque não cadastrado.", vbInformation, "Erro de Verificação!"
                    txt_valor_unitario.Text = ""
                    txt_valor_unitario.SetFocus
                    Exit Sub
                End If
            End If
            If txt_quantidade.Enabled Then
                txt_quantidade.SetFocus
            End If
        Else
            MsgBox "Produto não cadastrado.", vbInformation, "Atenção!"
            txt_produto.SetFocus
            Exit Sub
        End If
    Else
        txt_produto.Text = ""
        dtcboProduto.BoundText = 0
    End If
    
End Sub
Private Sub txt_quantidade_GotFocus()
    lOrigemFocus = "txt_quantidade"
    If lCodigoBarra Then
        If txt_produto.Text <> "" Then
            txt_quantidade_LostFocus
            GravaItem
            Exit Sub
        End If
    End If
    txt_quantidade.BackColor = &H800000
    txt_quantidade.ForeColor = &H80000005
    l_mensagem = Space(165) & "Informe a quantidade."
    If Val(txt_produto.Text) > 0 And (txt_quantidade.Text = "" Or txt_quantidade.Text = "1") Then
        If Produto.CodigoGrupo = lGrupoCombustivel Then
            txt_valor_total.SetFocus
            Exit Sub
        End If
    End If
    txt_quantidade.SelStart = 0
    txt_quantidade.SelLength = Len(txt_quantidade.Text)
End Sub
Private Sub txt_quantidade_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 3 Then
        KeyAscii = 0
        ChamaCalcLitros
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_valor_total.SetFocus
    End If
    Call ValidaValor(KeyAscii)
End Sub
Private Sub txt_quantidade_LostFocus()
    Dim i As Integer
    txt_quantidade.BackColor = &H80000005
    txt_quantidade.ForeColor = &H80000008
    txt_quantidade.Text = Format(txt_quantidade.Text, "###,##0.000")
    If g_string = "" Then
        txt_valor_total.Text = Format(Format(fValidaValor4(txt_valor_unitario.Text) * fValidaValor(txt_quantidade.Text), "###,##0.0000"), "###,##0.0000")
        i = Len(txt_valor_total.Text)
        txt_valor_total.Text = Mid(txt_valor_total.Text, 1, i - 2)
    Else
        g_string = ""
    End If
End Sub
Private Sub txt_senha_ponto_GotFocus()
    lOrigemFocus = "txt_senha_ponto"
    txt_senha_ponto.BackColor = &H800000
    txt_senha_ponto.ForeColor = &H80000005
    l_mensagem = Space(165) & "Informe a senha do funcionário."
    txt_senha_ponto.SelStart = 0
    txt_senha_ponto.SelLength = Len(txt_senha_ponto.Text)
End Sub
Private Sub txt_senha_ponto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok_ponto_Click
    End If
End Sub
Private Sub txt_senha_ponto_LostFocus()
    txt_senha_ponto.BackColor = &H80000005
    txt_senha_ponto.ForeColor = &H80000008
End Sub
Private Sub txt_telefone_GotFocus()
    l_mensagem = Space(165) & "Informe o telefone do cliente."
    txt_telefone.Text = fDesmascaraTelefone(txt_telefone.Text)
End Sub
Private Sub txt_telefone_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        
        If cmdIniciaProcessoNFCe.Enabled Then
            cmdIniciaProcessoNFCe.SetFocus
        End If
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_telefone_LostFocus()
    txt_telefone.Text = fMascaraTelefone(txt_telefone.Text)
End Sub
Private Sub txt_valor_desconto_GotFocus()
    l_mensagem = Space(165) & "Informe o valor do desconto."
    txt_valor_desconto.SelStart = 0
    txt_valor_desconto.SelLength = Len(txt_valor_desconto.Text)
End Sub
Private Sub txt_valor_desconto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_valor_recebido.SetFocus
    End If
    Call ValidaValor(KeyAscii)
End Sub
Private Sub txt_valor_desconto_LostFocus()
    txt_valor_desconto.Text = Format(txt_valor_desconto.Text, "###,##0.00")
End Sub
Private Sub txt_valor_recebido_GotFocus()
    l_mensagem = Space(165) & "Informe o valor recebido."
    lbl_valor_compra.Caption = Format(lTotalCupom - fValidaValor(txt_valor_desconto.Text), "###,##0.00")
    txt_valor_recebido.Text = Format(lTotalCupom - fValidaValor(txt_valor_desconto.Text), "###,##0.00")
    lbl_valor_troco.Caption = Format(0, "0.00")
    txt_valor_recebido.SelStart = 0
    txt_valor_recebido.SelLength = Len(txt_valor_recebido.Text)
    'If cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 4 Or cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 5 Then
    '    cmdIniciaProcessoNFCe.SetFocus
    'End If
End Sub
Private Sub txt_valor_recebido_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        If cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) < 2 Or cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) > 3 Then
            If cmdIniciaProcessoNFCe.Enabled Then
                cmdIniciaProcessoNFCe.SetFocus
            End If
        Else
            txt_numero_cheque.SetFocus
        End If
    End If
    Call ValidaValor(KeyAscii)
End Sub
Private Sub txt_valor_recebido_LostFocus()
    txt_valor_recebido.Text = Format(txt_valor_recebido.Text, "###,##0.00")
    lbl_valor_troco.Caption = Format(fValidaValor(txt_valor_recebido.Text) - fValidaValor(lbl_valor_compra.Caption), "###,##0.00")
End Sub
Private Sub txt_valor_total_GotFocus()
    lOrigemFocus = "txt_valor_total"
    If g_nivel_acesso > 1 Then
        If Produto.CodigoGrupo <> lGrupoCombustivel Then
            VerificaDescontoPersonalizado
            GravaItem
            Exit Sub
        End If
    End If
    txt_valor_total.BackColor = &H800000
    txt_valor_total.ForeColor = &H80000005
    l_mensagem = Space(165) & "Informe o valor da venda."
    txt_valor_total.SelStart = 0
    txt_valor_total.SelLength = Len(txt_valor_total.Text)
End Sub
Private Sub txt_valor_total_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        If Produto.CodigoGrupo = lGrupoCombustivel Then
            txt_valor_total.Text = Format(txt_valor_total.Text, "###,##0.00")
            If fValidaValor(txt_valor_total.Text) > 0 And fValidaValor(txt_valor_unitario.Text) > 0 Then
                txt_quantidade.Text = Format((fValidaValor(txt_valor_total.Text) / fValidaValor(txt_valor_unitario.Text)), "###,##0.000")
            Else
                txt_quantidade.Text = Format(0, "###,##0.000")
            End If
        End If
        KeyAscii = 0
        VerificaDescontoPersonalizado
        GravaItem
    End If
    Call ValidaValor(KeyAscii)
End Sub
Private Sub txt_valor_total_LostFocus()
    txt_valor_total.BackColor = &HC0C0C0
    txt_valor_total.ForeColor = &H80000008
    txt_valor_total.Text = Format(txt_valor_total.Text, "###,##0.00")
    
    
    If g_nome_empresa Like "*PELICANO*" Or g_nome_empresa Like "*BRITO BARROS*" Then
        If Estoque.Quantidade < Produto.EstoqueMinimo Then
            MsgBox "Este produto se encontra com estoque abaixo da quantidade minima.", vbInformation, "Atenção!"
        End If
    End If
End Sub
Private Sub txt_valor_unitario_GotFocus()
    lOrigemFocus = "txt_valor_unitario"
    If g_nivel_acesso > 1 Then
        txt_quantidade.SetFocus
        Exit Sub
    End If
End Sub
Private Sub txt_valor_unitario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_quantidade.SetFocus
    End If
    Call ValidaValor(KeyAscii)
End Sub
Private Sub txt_valor_unitario_LostFocus()
    txt_valor_unitario.Text = Format(txt_valor_unitario.Text, "###,##0.0000")
End Sub
Private Function BuscaNumeroCupom() As String
    Dim xString As String
    Dim xRetorno As Long
    Dim xData As String
    Dim xHora As String
    Dim NumeroArquivo As Integer
    
    On Error GoTo FileError
    
    
    BuscaNumeroCupom = "OK"
    
    xData = Format(Now, "dd/mm/yyyy") 'ALEX - NFCe 'Space(6)
    xHora = Format(Now, "HH:mm:ss")  'ALEX - NFCe 'Space(6)
    
    lData = CDate(xData) 'ALEX - NFCe 'CDate(Mid(xData, 1, 2) & "/" & Mid(xData, 3, 2) & "/20" & Mid(xData, 5, 2))
    lDataCupom = lData
    lHora = CDate(xHora) 'ALEX - NFCe 'Format(Mid(xHora, 1, 2), "00") & ":" & Format(Mid(xHora, 3, 2), "00") & ":" & Format(Mid(xHora, 5, 2), "00")
    If l_flag_cupom_fiscal = "F" Then
'        If MovCupomFiscal.LocalizarUltimo(g_empresa, lNumeroPDV) Then
'            lNumeroCupom = MovCupomFiscal.NumeroCupom + 1
'        End If
        xString = ConfiguracaoDiversa.BuscaProximoCodigo(g_empresa, "NFCe: Numero", True)
        If Len(xString) > 0 Then
            lNumeroCupom = CLng(RetiraString(1, xString))
            lSerieNFCe = RetiraString(2, xString)
        Else
            lNumeroCupom = 1
            lSerieNFCe = "1"
        End If
        lOrdem = 1
    Else
'        If MovCupomFiscal.LocalizarUltimo(g_empresa, lNumeroPDV) Then
'            lNumeroCupom = MovCupomFiscal.NumeroCupom
'            lOrdem = MovCupomFiscal.Ordem + 1
'        End If
        lOrdem = lOrdem + 1
    End If
    'lHoraPegouNumeroCupom = Now
    'Call VerificaSeExisteCupom
    Exit Function
FileError:
    MsgBox "Não foi possível criar o novo cupom fiscal.", vbCritical, "Erro Grave!"
    Exit Function
End Function
Private Sub BuscaNumeroDeSerie()
    Dim x_string As String
    Dim NumeroArquivo As Integer
    On Error GoTo FileError
    'If lExisteImpressora Then
    '    'busca número de série
    '    Call Abre_ProtocoloCF(1)
    '    ComandoCF = Chr(27) + "|35|00|" + Chr(27)
    '    Envia_ComandoCF
    '    Fecha_ProtocoloCF
    '    NumeroArquivo = FreeFile
    '    Open "MP20FI.RET" For Input As NumeroArquivo
    '    Input #NumeroArquivo, x_string
    '    Close NumeroArquivo
    '    If g_nome_empresa = "T-Kar Posto Shopping Ltda" And x_string = "4708990404338" Then
    '        Exit Sub
    '    End If
    '    MsgBox "Número de Série da Impressora Fiscal ->" & x_string & "<-" & Chr(13) & "Empresa ->" & Trim(g_empresa) & "<-", vbInformation, "Número de Série"
    '    MsgBox "O sistema será finalizado", vbCritical, "Erro Interno Fatal"
    '    End
    'End If
    Exit Sub
FileError:
    MsgBox "Não foi possível verificar o número de série.", vbCritical, "Erro Grave!"
    Exit Sub
End Sub
Private Sub txtQuantidadeDescarregamento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        chkDesconto.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
