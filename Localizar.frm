VERSION 5.00
Begin VB.Form localizar 
   Caption         =   "Localizar"
   ClientHeight    =   1125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5280
   Icon            =   "Localizar.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1125
   ScaleWidth      =   5280
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd_cancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   300
      Width           =   1335
   End
   Begin VB.CheckBox chk_diferenciar 
      Caption         =   "&Diferenciar Maiúsculo/Minúsculo"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2955
   End
   Begin VB.TextBox txt_localizar 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Localizar por:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   945
   End
End
Attribute VB_Name = "localizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_cancelar_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    txt_localizar.Text = g_string
    g_string = ""
End Sub
Private Sub txt_localizar_GotFocus()
    txt_localizar.SelStart = 0
    txt_localizar.SelLength = Len(txt_localizar.Text)
End Sub

Private Sub txt_localizar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        g_string = txt_localizar.Text & "|@|"
        If chk_diferenciar.Value = True Then
            g_string = g_string & "True" & "|@|"
        Else
            g_string = g_string & "False" & "|@|"
        End If
        Unload Me
    End If
End Sub
