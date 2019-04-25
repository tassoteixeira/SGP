VERSION 5.00
Begin VB.Form frm_sobre2 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3060
   ClientLeft      =   2775
   ClientTop       =   2385
   ClientWidth     =   5040
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3060
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   120
      Top             =   1320
   End
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin VB.CommandButton cmd_ok 
         Caption         =   "&OK"
         Height          =   375
         Left            =   3840
         TabIndex        =   1
         Top             =   2220
         Width           =   615
      End
      Begin VB.Timer Timer2 
         Interval        =   100
         Left            =   0
         Top             =   1620
      End
      Begin VB.PictureBox Picture1 
         Height          =   675
         Left            =   1980
         Picture         =   "frm_sobre2.frx":0000
         ScaleHeight     =   615
         ScaleWidth      =   2670
         TabIndex        =   5
         Top             =   1320
         Width           =   2730
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Cupom Fiscal - Cerrado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   2220
         TabIndex        =   6
         Top             =   480
         Width           =   2355
      End
      Begin VB.Image Image2 
         Height          =   255
         Left            =   4500
         Top             =   60
         Width           =   315
      End
      Begin VB.Label lbl_empresa 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "T-Kar Posto Shopping Ltda."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   2280
         Width           =   3135
      End
      Begin VB.Label Label3 
         Caption         =   "Cópia Licenciada Para:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "P/ Windows-95/98/NT  Versão 1.1"
         Height          =   255
         Left            =   2160
         TabIndex        =   2
         Top             =   1080
         Width           =   2595
      End
      Begin VB.Image Image1 
         DragMode        =   1  'Automatic
         Height          =   840
         Left            =   480
         Picture         =   "frm_sobre2.frx":4B146
         Top             =   480
         Width           =   825
      End
   End
End
Attribute VB_Name = "frm_sobre2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmd_ok_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    Screen.MousePointer = 1
    CentraForm Me
    lbl_empresa = g_nome_empresa
End Sub
Private Sub Image2_Click()
    g_lmc = g_lmc + 1
End Sub
Private Sub Timer1_Timer()
    Static flag As Integer
    If flag = 0 Then
        Image1.Picture = LoadPicture("\vb5\sgp\icons\bomba1.bmp")
        flag = flag + 1
    ElseIf flag = 1 Then
        Image1.Picture = LoadPicture("\vb5\sgp\icons\bomba2.bmp")
        flag = flag + 1
    ElseIf flag = 2 Then
        Image1.Picture = LoadPicture("\vb5\sgp\icons\bomba3.bmp")
        flag = flag + 1
    ElseIf flag = 3 Then
        Image1.Picture = LoadPicture("\vb5\sgp\icons\bomba4.bmp")
        flag = flag + 1
    ElseIf flag = 4 Then
        Image1.Picture = LoadPicture("\vb5\sgp\icons\bomba5.bmp")
        flag = 0
    End If
End Sub
Private Sub Timer2_Timer()
    Static flag_2 As Integer
    flag_2 = flag_2 + 1
'    flag_2 = 1
    Picture1.Cls
    Picture1.CurrentX = 420
    Picture1.CurrentY = 600 - (25 * flag_2)
    Picture1.Print "Cerrado Informática Ltda."
    Picture1.CurrentX = 325
    Picture1.Print "Desenvolvendo Tecnologia."
    Picture1.Print ""
    Picture1.CurrentX = 525
    Picture1.Print "Autor:  Tasso Teixeira"
    Picture1.CurrentX = 530
    Picture1.Print "Fone: 0xx62 941-3044"
    Picture1.Print ""
    Picture1.CurrentX = 600
    Picture1.Print "Rogério R. Bailona"
    Picture1.CurrentX = 530
    Picture1.Print "Fone: 0xx62 991-9521"
    If flag_2 = 85 Then
        flag_2 = 0
    End If
End Sub
