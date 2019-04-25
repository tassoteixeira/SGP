VERSION 5.00
Begin VB.Form splash2 
   Caption         =   "Cupom Fiscal - Cerrado"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8385
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   8385
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt_mensagem 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1860
      TabIndex        =   1
      Top             =   6240
      Width           =   4575
   End
   Begin VB.Timer Timer2 
      Interval        =   25
      Left            =   0
      Top             =   1260
   End
   Begin VB.Timer Timer1 
      Interval        =   900
      Left            =   0
      Top             =   2160
   End
   Begin VB.PictureBox Picture1 
      Height          =   6555
      Left            =   0
      Picture         =   "splash2.frx":0000
      ScaleHeight     =   6495
      ScaleWidth      =   8295
      TabIndex        =   0
      Top             =   0
      Width           =   8355
      Begin VB.Label lbl_mensagem 
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1800
         TabIndex        =   3
         Top             =   5760
         Visible         =   0   'False
         Width           =   4635
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cerrado Informática."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   555
         Left            =   1560
         TabIndex        =   2
         Top             =   240
         Width           =   5475
      End
   End
End
Attribute VB_Name = "splash2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x_tempo As Integer
Dim x_tempo2 As Integer
Dim x_flag As Boolean
Private Sub Form_Load()
    CentraForm Me
    x_tempo = 0
    x_tempo2 = 0
    x_flag = False
End Sub
Private Sub Timer1_Timer()
    x_tempo = x_tempo + 1
    If x_tempo = 1 Then
        Call ChamaDrive
'
'        ChDrive "C"
'
        ChDir "\VB5\SGP\DATA"
    ElseIf x_tempo = 2 Then
        Set bd_sgp = OpenDatabase("SGP_DATA.MDB")
    ElseIf x_tempo = 3 Then
        Set bd_sgp_b = OpenDatabase("SGP_DATA_BAIXA.MDB")
    ElseIf x_tempo = 4 Then
        Set bd_sgp_m = OpenDatabase("SGP_DATA_MOVIMENTO.MDB")
    ElseIf x_tempo = 5 Then
        Unload Me
    End If
End Sub
Private Sub Timer2_Timer()
    Dim x_mensagem As String
    x_mensagem = Space(90) & "Aguarde! Abrindo os bancos de dados..."
    x_tempo2 = x_tempo2 + 1
    If x_tempo2 = 1 Then
        lbl_mensagem = x_mensagem
        txt_mensagem = x_mensagem
    Else
        If x_tempo2 <= Len(x_mensagem) Then
            lbl_mensagem = Mid(x_mensagem, x_tempo2, Len(x_mensagem) - x_tempo2)
            txt_mensagem = Mid(x_mensagem, x_tempo2, Len(x_mensagem) - x_tempo2)
        Else
            x_tempo2 = 0
        End If
    End If
End Sub
