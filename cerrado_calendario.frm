VERSION 5.00
Begin VB.Form cerrado_calendario 
   Caption         =   "Calendário Cerrado"
   ClientHeight    =   4905
   ClientLeft      =   1350
   ClientTop       =   1680
   ClientWidth     =   4395
   Icon            =   "cerrado_calendario.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4905
   ScaleWidth      =   4395
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Height          =   555
      Left            =   3720
      Picture         =   "cerrado_calendario.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   63
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   4260
      Width           =   555
   End
   Begin VB.CommandButton cmd_data_hoje 
      Caption         =   "Data Hoje"
      Height          =   315
      Left            =   120
      TabIndex        =   62
      ToolTipText     =   "Posiciona na data de hoje."
      Top             =   4500
      Width           =   3495
   End
   Begin VB.TextBox txt_diferenca 
      Height          =   315
      Left            =   1740
      MaxLength       =   5
      TabIndex        =   58
      Top             =   3840
      Width           =   915
   End
   Begin VB.CommandButton cmd_data_final 
      Caption         =   "Data Final"
      Height          =   315
      Left            =   3120
      Picture         =   "cerrado_calendario.frx":171C
      Style           =   1  'Graphical
      TabIndex        =   60
      ToolTipText     =   "Seleciona data final."
      Top             =   3840
      Width           =   1155
   End
   Begin VB.CommandButton cmd_data_inicial 
      Caption         =   "Data Inicial"
      Height          =   315
      Left            =   120
      Picture         =   "cerrado_calendario.frx":382E
      Style           =   1  'Graphical
      TabIndex        =   56
      ToolTipText     =   "Seleciona data inicial."
      Top             =   3840
      Width           =   1155
   End
   Begin VB.Frame frm_dados 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   4155
      Begin VB.CommandButton cmd_ter 
         Appearance      =   0  'Flat
         Caption         =   "Ter"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1260
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   840
         Width           =   495
      End
      Begin VB.CommandButton cmd_qua 
         Appearance      =   0  'Flat
         Caption         =   "Qua"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   840
         Width           =   495
      End
      Begin VB.CommandButton cmd_qui 
         Appearance      =   0  'Flat
         Caption         =   "Qui"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2340
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   840
         Width           =   495
      End
      Begin VB.CommandButton cmd_sex 
         Appearance      =   0  'Flat
         Caption         =   "Sex"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   840
         Width           =   495
      End
      Begin VB.CommandButton cmd_sab 
         Appearance      =   0  'Flat
         Caption         =   "Sab"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3420
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   840
         Width           =   495
      End
      Begin VB.CommandButton cmd_seg 
         Appearance      =   0  'Flat
         Caption         =   "Seg"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   840
         Width           =   495
      End
      Begin VB.CommandButton cmd_dom 
         Appearance      =   0  'Flat
         Caption         =   "Dom"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   840
         Width           =   495
      End
      Begin VB.ComboBox cbo_ano 
         Height          =   330
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Seleciona o ano."
         Top             =   420
         Width           =   1335
      End
      Begin VB.ComboBox cbo_mes 
         Height          =   330
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Seleciona o mês."
         Top             =   420
         Width           =   2175
      End
      Begin VB.Frame frmFrame1
         Height          =   2535
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   3855
         BackColor       =   16777215
         Begin VB.CommandButton cmd_dia 
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   42
            Left            =   3360
            Picture         =   "cerrado_calendario.frx":5940
            Style           =   1  'Graphical
            TabIndex        =   54
            Top             =   2160
            Width           =   375
         End
         Begin VB.CommandButton cmd_dia 
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   41
            Left            =   2820
            Style           =   1  'Graphical
            TabIndex        =   53
            Top             =   2160
            Width           =   375
         End
         Begin VB.CommandButton cmd_dia 
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   40
            Left            =   2280
            Style           =   1  'Graphical
            TabIndex        =   52
            Top             =   2160
            Width           =   375
         End
         Begin VB.CommandButton cmd_dia 
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   39
            Left            =   1740
            Style           =   1  'Graphical
            TabIndex        =   51
            Top             =   2160
            Width           =   375
         End
         Begin VB.CommandButton cmd_dia 
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   38
            Left            =   1200
            Style           =   1  'Graphical
            TabIndex        =   50
            Top             =   2160
            Width           =   375
         End
         Begin VB.CommandButton cmd_dia 
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   37
            Left            =   660
            Style           =   1  'Graphical
            TabIndex        =   49
            Top             =   2160
            Width           =   375
         End
         Begin VB.CommandButton cmd_dia 
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   36
            Left            =   120
            Picture         =   "cerrado_calendario.frx":6C1A
            Style           =   1  'Graphical
            TabIndex        =   48
            Top             =   2160
            Width           =   375
         End
         Begin VB.CommandButton cmd_dia 
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   35
            Left            =   3360
            Picture         =   "cerrado_calendario.frx":7EF4
            Style           =   1  'Graphical
            TabIndex        =   47
            Top             =   1800
            Width           =   375
         End
         Begin VB.CommandButton cmd_dia 
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   34
            Left            =   2820
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   1800
            Width           =   375
         End
         Begin VB.CommandButton cmd_dia 
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   33
            Left            =   2280
            Style           =   1  'Graphical
            TabIndex        =   45
            Top             =   1800
            Width           =   375
         End
         Begin VB.CommandButton cmd_dia 
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   32
            Left            =   1740
            Style           =   1  'Graphical
            TabIndex        =   44
            Top             =   1800
            Width           =   375
         End
         Begin VB.CommandButton cmd_dia 
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   31
            Left            =   1200
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   1800
            Width           =   375
         End
         Begin VB.CommandButton cmd_dia 
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   30
            Left            =   660
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   1800
            Width           =   375
         End
         Begin VB.CommandButton cmd_dia 
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   29
            Left            =   120
            Picture         =   "cerrado_calendario.frx":91CE
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   1800
            Width           =   375
         End
         Begin VB.CommandButton cmd_dia 
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   28
            Left            =   3360
            Picture         =   "cerrado_calendario.frx":A4A8
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   1440
            Width           =   375
         End
         Begin VB.CommandButton cmd_dia 
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   27
            Left            =   2820
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   1440
            Width           =   375
         End
         Begin VB.CommandButton cmd_dia 
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   26
            Left            =   2280
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   1440
            Width           =   375
         End
         Begin VB.CommandButton cmd_dia 
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   25
            Left            =   1740
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   1440
            Width           =   375
         End
         Begin VB.CommandButton cmd_dia 
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   24
            Left            =   1200
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   1440
            Width           =   375
         End
         Begin VB.CommandButton cmd_dia 
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   23
            Left            =   660
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   1440
            Width           =   375
         End
         Begin VB.CommandButton cmd_dia 
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   22
            Left            =   120
            Picture         =   "cerrado_calendario.frx":B782
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   1440
            Width           =   375
         End
         Begin VB.CommandButton cmd_dia 
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   21
            Left            =   3360
            Picture         =   "cerrado_calendario.frx":CA5C
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   1080
            Width           =   375
         End
         Begin VB.CommandButton cmd_dia 
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   20
            Left            =   2820
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   1080
            Width           =   375
         End
         Begin VB.CommandButton cmd_dia 
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   19
            Left            =   2280
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   1080
            Width           =   375
         End
         Begin VB.CommandButton cmd_dia 
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   18
            Left            =   1740
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   1080
            Width           =   375
         End
         Begin VB.CommandButton cmd_dia 
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   17
            Left            =   1200
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   1080
            Width           =   375
         End
         Begin VB.CommandButton cmd_dia 
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   16
            Left            =   660
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   1080
            Width           =   375
         End
         Begin VB.CommandButton cmd_dia 
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   15
            Left            =   120
            Picture         =   "cerrado_calendario.frx":DD36
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   1080
            Width           =   375
         End
         Begin VB.CommandButton cmd_dia 
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   14
            Left            =   3360
            Picture         =   "cerrado_calendario.frx":F010
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   720
            Width           =   375
         End
         Begin VB.CommandButton cmd_dia 
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   13
            Left            =   2820
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   720
            Width           =   375
         End
         Begin VB.CommandButton cmd_dia 
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   12
            Left            =   2280
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   720
            Width           =   375
         End
         Begin VB.CommandButton cmd_dia 
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   11
            Left            =   1740
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   720
            Width           =   375
         End
         Begin VB.CommandButton cmd_dia 
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   10
            Left            =   1200
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   720
            Width           =   375
         End
         Begin VB.CommandButton cmd_dia 
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   9
            Left            =   660
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   720
            Width           =   375
         End
         Begin VB.CommandButton cmd_dia 
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   8
            Left            =   120
            Picture         =   "cerrado_calendario.frx":102EA
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   720
            Width           =   375
         End
         Begin VB.CommandButton cmd_dia 
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   7
            Left            =   3360
            Picture         =   "cerrado_calendario.frx":115C4
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   360
            Width           =   375
         End
         Begin VB.CommandButton cmd_dia 
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   6
            Left            =   2820
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   360
            Width           =   375
         End
         Begin VB.CommandButton cmd_dia 
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   5
            Left            =   2280
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   360
            Width           =   375
         End
         Begin VB.CommandButton cmd_dia 
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   1740
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   360
            Width           =   375
         End
         Begin VB.CommandButton cmd_dia 
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   1200
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   360
            Width           =   375
         End
         Begin VB.CommandButton cmd_dia 
            Appearance      =   0  'Flat
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   660
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   360
            Width           =   375
         End
         Begin VB.CommandButton cmd_dia 
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   120
            Picture         =   "cerrado_calendario.frx":1289E
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   360
            Width           =   375
         End
      End
      Begin VB.Label Label3 
         Caption         =   "&Ano"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   3
         Top             =   180
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "&Mês"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   180
         Width           =   615
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Data de &Hoje"
      Height          =   255
      Left            =   180
      TabIndex        =   61
      Top             =   4260
      Width           =   975
   End
   Begin VB.Label lbl_diferenca 
      Caption         =   "&Diferença"
      Height          =   255
      Left            =   1740
      TabIndex        =   57
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label lbl_data_inicial 
      Caption         =   "Data &Inicial"
      Height          =   255
      Left            =   180
      TabIndex        =   55
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label lbl_data_final 
      Caption         =   "Data &Final"
      Height          =   255
      Left            =   3180
      TabIndex        =   59
      Top             =   3600
      Width           =   975
   End
End
Attribute VB_Name = "cerrado_calendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_tipo_data As String * 1
Dim l_transporta_data As Boolean
Dim l_data_i As Date
Dim l_data_f As Date
Dim l_qtd_click As Integer
Private Sub AtualizaMesAno(x_mes As Integer, x_ano As Integer)
    Dim i As Integer
    For i = 0 To cbo_mes.ListCount - 1
        If cbo_mes.ItemData(i) = x_mes Then
            cbo_mes.ListIndex = i
            Exit For
        End If
    Next
    For i = 0 To cbo_ano.ListCount - 1
        If cbo_ano.ItemData(i) = x_ano Then
            cbo_ano.ListIndex = i
            Exit For
        End If
    Next
End Sub
Private Sub CalculaDiferenca()
    txt_diferenca = DateDiff("d", CDate(cmd_data_inicial.Caption), CDate(cmd_data_final.Caption))
End Sub
Private Sub Finaliza()
    'tbl_funcionario.Close
End Sub
Private Sub LimpaDias()
    Dim i As Integer
    For i = 1 To 42
        cmd_dia(i).Caption = ""
        cmd_dia(i).Visible = True
    Next
End Sub
Private Sub MontaCalendario()
    Dim i As Integer
    Dim x_data As Date
    Dim x_data_teste As Date
    x_data = CDate(Format(Date, "dd") & "/" & Format(cbo_mes.ItemData(cbo_mes.ListIndex), "00") & "/" & Format(cbo_ano.ItemData(cbo_ano.ListIndex), "0000"))
    x_data_teste = CDate("01/" & Month(x_data) & "/" & Year(x_data))
    For i = 1 To 42
        cmd_dia(i).Caption = ""
        cmd_dia(i).Visible = True
        cmd_dia(i).Picture = LoadPicture()
        If i = 1 Or i = 8 Or i = 15 Or i = 22 Or i = 29 Or i = 36 Then
            cmd_dia(i).Picture = LoadPicture("\vb5\sgp\icons\dia_vermelho.bmp")
        End If
        If i = 7 Or i = 14 Or i = 21 Or i = 28 Or i = 35 Or i = 42 Then
            cmd_dia(i).Picture = LoadPicture("\vb5\sgp\icons\dia_azul.bmp")
        End If
    Next
    i = Format(x_data_teste, "w") - 1
    Do Until Month(x_data_teste) <> Month(x_data)
        i = i + 1
        cmd_dia(i).Caption = Day(x_data_teste)
        cmd_dia(i).Enabled = True
        x_data_teste = x_data_teste + 1
    Loop
    For i = 1 To 42
        If cmd_dia(i).Caption = "" Then
            cmd_dia(i).Visible = False
        End If
        If Val(Format(cmd_data_inicial.Caption, "mm")) = cbo_mes.ItemData(cbo_mes.ListIndex) And Val(Format(cmd_data_inicial.Caption, "yyyy")) = cbo_ano.ItemData(cbo_ano.ListIndex) Then
            If Val(cmd_dia(i).Caption) = Val(Format(cmd_data_inicial.Caption, "dd")) Then
                cmd_dia(i).Picture = LoadPicture("\vb5\sgp\icons\dia_inicial.bmp")
            End If
        End If
        If Val(Format(cmd_data_final.Caption, "mm")) = cbo_mes.ItemData(cbo_mes.ListIndex) And Val(Format(cmd_data_final.Caption, "yyyy")) = cbo_ano.ItemData(cbo_ano.ListIndex) Then
            If Val(cmd_dia(i).Caption) = Val(Format(cmd_data_final.Caption, "dd")) Then
                cmd_dia(i).Picture = LoadPicture("\vb5\sgp\icons\dia_final.bmp")
            End If
        End If
    Next
End Sub
Private Sub MudaMensagem(x_tipo_data As String)
    Dim i As Integer
    Dim x_mensagem As String
    If x_tipo_data = "f" Then
        l_tipo_data = "f"
        x_mensagem = "Marca a data final."
    ElseIf x_tipo_data = "i" Then
        l_tipo_data = "i"
        x_mensagem = "Marca a data inicial."
    End If
    For i = 1 To 42
'        cmd_dia(i).Visible = True
        cmd_dia(i).ToolTipText = x_mensagem
    Next
End Sub
Private Sub PreencheCboAno()
    Dim i As Integer
    cbo_ano.Clear
    For i = 1900 To 2030
        cbo_ano.AddItem i
        cbo_ano.ItemData(cbo_ano.NewIndex) = i
    Next
End Sub
Private Sub PreencheCboMes()
    Dim i As Integer
    Dim x_data As Date
    cbo_mes.Clear
    x_data = "01/01/2000"
    For i = 1 To 12
        x_data = CDate(Mid(x_data, 1, 3) & Format(i, "00") & Mid(x_data, 6, 5))
        cbo_mes.AddItem Format(x_data, "mmmm")
        cbo_mes.ItemData(cbo_mes.NewIndex) = i
    Next
End Sub
Private Sub TransportaData()
    l_qtd_click = l_qtd_click + 1
    If l_transporta_data Then
        g_string = l_data_i & "|@|"
        If l_qtd_click > 1 Then
            If l_data_i <> l_data_f Then
                g_string = g_string & l_data_f & "|@|"
            End If
        End If
    End If
End Sub
Private Sub cbo_ano_KeyPress(KeyAscii As Integer)
    MontaCalendario
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_ano.ListIndex = cbo_ano.ListIndex
        cbo_mes.SetFocus
    End If
End Sub
Private Sub cbo_ano_KeyUp(KeyCode As Integer, Shift As Integer)
    MontaCalendario
End Sub
Private Sub cbo_ano_LostFocus()
'    MontaCalendario
End Sub
Private Sub cbo_mes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_ano.ListIndex = cbo_ano.ListIndex
        cbo_ano.SetFocus
    End If
End Sub
Private Sub cbo_mes_KeyUp(KeyCode As Integer, Shift As Integer)
    MontaCalendario
End Sub
Private Sub cbo_mes_LostFocus()
'    MontaCalendario
End Sub
Private Sub cmd_data_final_Click()
    Call MudaMensagem("f")
End Sub
Private Sub cmd_data_inicial_Click()
    Call MudaMensagem("i")
End Sub
Private Sub cmd_dia_Click(Index As Integer)
    If l_tipo_data = "i" Then
        cmd_data_inicial.Caption = CDate(Format(cmd_dia(Index).Caption, "00") & "/" & Format(cbo_mes.ItemData(cbo_mes.ListIndex), "00") & "/" & Format(cbo_ano.ItemData(cbo_ano.ListIndex), "0000"))
        l_data_i = cmd_data_inicial.Caption
        Call MudaMensagem("f")
        cmd_data_final.SetFocus
    Else
        cmd_data_final.Caption = CDate(Format(cmd_dia(Index).Caption, "00") & "/" & Format(cbo_mes.ItemData(cbo_mes.ListIndex), "00") & "/" & Format(cbo_ano.ItemData(cbo_ano.ListIndex), "0000"))
        l_data_f = cmd_data_final.Caption
        Call MudaMensagem("i")
        cmd_data_inicial.SetFocus
    End If
    TransportaData
    MontaCalendario
    CalculaDiferenca
End Sub
Private Sub cmd_data_hoje_Click()
    cmd_data_hoje.Caption = Format(Date, "dddd, d") & " de " & Format(Date, "mmmm") & " de " & Format(Date, "yyyy")
    Call AtualizaMesAno(Format(Date, "MM"), Format(Date, "yyyy"))
    MontaCalendario
End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub Form_Activate()
    Screen.MousePointer = 1
    LimpaDias
    Call AtualizaMesAno(Format(Date, "MM"), Format(Date, "yyyy"))
    MontaCalendario
End Sub
Private Sub Form_Load()
    CentraForm Me
    PreencheCboMes
    PreencheCboAno
    cmd_data_inicial.Caption = Format(Date, "dd/mm/yyyy")
    cmd_data_final.Caption = Format(Date, "dd/mm/yyyy")
    cmd_data_hoje.Caption = Format(Date, "dddd, d") & " de " & Format(Date, "mmmm") & " de " & Format(Date, "yyyy")
    Call MudaMensagem("i")
    l_data_i = cmd_data_inicial.Caption
    l_data_f = cmd_data_final.Caption
    l_transporta_data = False
    'If IsDate(g_string) Then
        g_string = g_string & "|@|"
        l_transporta_data = True
    'End If
    l_qtd_click = 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub txt_diferenca_GotFocus()
    txt_diferenca.SelStart = 0
    txt_diferenca.SelLength = Len(txt_diferenca)
End Sub
Private Sub txt_diferenca_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_data_final.SetFocus
        Call MudaMensagem("f")
    End If
End Sub
Private Sub txt_diferenca_LostFocus()
    cmd_data_final.Caption = CDate(cmd_data_inicial.Caption) + Val(txt_diferenca)
    l_data_f = cmd_data_final.Caption
    TransportaData
    MontaCalendario
End Sub
