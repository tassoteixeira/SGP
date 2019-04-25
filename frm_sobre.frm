VERSION 5.00
Begin VB.Form frm_sobre 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3090
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
   ScaleHeight     =   3090
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   120
      Top             =   480
   End
   Begin VB.Frame Frame1 
      Height          =   2835
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
         Top             =   780
      End
      Begin VB.PictureBox Picture1 
         Height          =   675
         Left            =   1980
         Picture         =   "frm_sobre.frx":0000
         ScaleHeight     =   615
         ScaleWidth      =   2670
         TabIndex        =   5
         Top             =   1320
         Width           =   2730
      End
      Begin VB.Label Label6 
         Caption         =   "N. Série:"
         Height          =   255
         Left            =   180
         TabIndex        =   9
         Top             =   2460
         Width           =   615
      End
      Begin VB.Label lbl_numero_serie 
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Número de Série"
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
         Left            =   960
         TabIndex        =   8
         Top             =   2460
         Width           =   1485
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Postos de Combustíveis"
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
         Left            =   1320
         TabIndex        =   7
         Top             =   600
         Width           =   2955
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Sistema Gerênciador de"
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
         Left            =   1320
         TabIndex        =   6
         Top             =   300
         Width           =   2955
      End
      Begin VB.Image Image2 
         Height          =   255
         Left            =   4500
         Top             =   60
         Width           =   315
      End
      Begin VB.Label lbl_empresa 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cerrado Informática"
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
         Left            =   180
         TabIndex        =   4
         Top             =   2100
         Width           =   3135
      End
      Begin VB.Label Label3 
         Caption         =   "Cópia Licenciada Para:"
         Height          =   255
         Left            =   180
         TabIndex        =   3
         Top             =   1860
         Width           =   1695
      End
      Begin VB.Label lblVersao 
         Alignment       =   2  'Center
         Caption         =   "Versão x.x.x para Windows XP ao 7"
         Height          =   255
         Left            =   1980
         TabIndex        =   2
         Top             =   1080
         Width           =   2715
      End
      Begin VB.Image Image1 
         DragMode        =   1  'Automatic
         Height          =   840
         Left            =   480
         Picture         =   "frm_sobre.frx":4B146
         Top             =   480
         Width           =   825
      End
   End
End
Attribute VB_Name = "frm_sobre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmd_ok_Click()
    Unload Me
End Sub
Private Sub cmd_ok_KeyPress(KeyAscii As Integer)
    Dim xString As String
    'Ctrl + l
    If KeyAscii = 4 Then
        xString = "Data Demonstração:" & gDataDemonstracao
        MsgBox xString
    'Ctrl + d
    ElseIf KeyAscii = 12 Then
        xString = "Data Limite:" & gDataLimiteUso
        MsgBox xString
    'Ctrl + t
    ElseIf KeyAscii = 20 Then
        xString = "Data Limite:" & gDataLimiteUso & Chr(10) & "Data Demonstração:" & gDataDemonstracao
        If VerificaHD Then
            xString = "Registro OK" & Chr(10) & xString
            MsgBox xString
        Else
            xString = "Falha de Registro" & Chr(10) & xString
            MsgBox xString
        End If
    End If
End Sub
Private Sub Form_Load()
Dim lNumeroHd As String
    Call GravaAuditoria(1, Me.name, 1, "")
    lblVersao.Caption = "Versão " & gVersaoSGP & " para Windows XP ao 7"
    Screen.MousePointer = 1
    CentraForm Me
    lbl_empresa = g_nome_empresa
    lNumeroHd = DriveSerial(Left("C:", 1))
    lbl_numero_serie.Caption = Chr(34) & "2000-0037-12-01-" & lNumeroHd & Chr(34)
End Sub
Private Sub Image2_Click()
    g_lmc = g_lmc + 1
    Call GravaAuditoria(1, Me.name, 10, g_lmc)
End Sub
Private Sub Timer1_Timer()
    Static Flag As Integer
    If Flag = 0 Then
        Image1.Picture = LoadPicture("\vb5\sgp\icons\bomba1.bmp")
        Flag = Flag + 1
    ElseIf Flag = 1 Then
        Image1.Picture = LoadPicture("\vb5\sgp\icons\bomba2.bmp")
        Flag = Flag + 1
    ElseIf Flag = 2 Then
        Image1.Picture = LoadPicture("\vb5\sgp\icons\bomba3.bmp")
        Flag = Flag + 1
    ElseIf Flag = 3 Then
        Image1.Picture = LoadPicture("\vb5\sgp\icons\bomba4.bmp")
        Flag = Flag + 1
    ElseIf Flag = 4 Then
        Image1.Picture = LoadPicture("\vb5\sgp\icons\bomba5.bmp")
        Flag = 0
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
    'Picture1.CurrentX = 685
    'Picture1.Print "Fone: 8414-9593"
    Picture1.CurrentX = 685
    Picture1.Print "Fone: 8436-4444"
    If flag_2 = 65 Then
        flag_2 = 0
    End If
End Sub
