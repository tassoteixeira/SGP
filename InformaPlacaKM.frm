VERSION 5.00
Begin VB.Form InformaPlacaKM 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Informa Placa e KM de Veículo"
   ClientHeight    =   2280
   ClientLeft      =   3330
   ClientTop       =   2790
   ClientWidth     =   3795
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "InformaPlacaKM.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Bancos"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2280
   ScaleWidth      =   3795
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Picture         =   "InformaPlacaKM.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Confirma o registro atual."
      Top             =   1320
      Width           =   795
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2880
      Picture         =   "InformaPlacaKM.frx":1914
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Cancela o registro atual."
      Top             =   1320
      Width           =   795
   End
   Begin VB.Frame frmDados 
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3585
      Begin VB.TextBox txtPlacaNumero 
         Height          =   285
         Left            =   2940
         MaxLength       =   4
         TabIndex        =   3
         Top             =   240
         Width           =   555
      End
      Begin VB.TextBox txtPlacaLetra 
         Height          =   285
         Left            =   2340
         MaxLength       =   3
         TabIndex        =   2
         Top             =   240
         Width           =   555
      End
      Begin VB.TextBox txtKM 
         Height          =   285
         Left            =   2340
         MaxLength       =   10
         TabIndex        =   5
         Top             =   660
         Width           =   1095
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Placa"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "&KM"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   660
         Width           =   2175
      End
   End
End
Attribute VB_Name = "InformaPlacaKM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmd_cancelar_Click()
    g_string = ""
    Unload Me
End Sub
Private Sub cmd_ok_Click()
    g_string = txtPlacaLetra.Text & "|@|"
    g_string = g_string & txtPlacaNumero.Text & "|@|"
    g_string = g_string & txtKM.Text & "|@|"
    Unload Me
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF11 Then
        KeyCode = 0
        cmd_ok_Click
    ElseIf KeyCode = vbKeyF12 Then
        KeyCode = 0
        cmd_cancelar_Click
    End If
End Sub
Private Sub Form_Load()
    CentraForm Me
    Screen.MousePointer = 1
    txtPlacaLetra.Text = ""
    txtPlacaNumero.Text = ""
    txtKM.Text = ""
End Sub
Private Sub txtKM_GotFocus()
    txtKM.SelStart = 0
    txtKM.SelLength = Len(txtKM.Text)
End Sub
Private Sub txtKM_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txtKM_LostFocus()
    If Val(txtKM.Text) > 0 Then
        txtKM.Text = Format(CLng(txtKM.Text), "###,###,##0")
    End If
End Sub
Private Sub txtPlacaLetra_GotFocus()
    txtPlacaLetra.SelStart = 0
    txtPlacaLetra.SelLength = Len(txtPlacaLetra.Text)
End Sub
Private Sub txtPlacaLetra_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtPlacaNumero.SetFocus
    End If
    Call ValidaLetra(KeyAscii)
    If Len(txtPlacaLetra.Text) = 3 Then
        txtPlacaNumero.SetFocus
    End If
End Sub
Private Sub txtPlacaNumero_GotFocus()
    txtPlacaNumero.SelStart = 0
    txtPlacaNumero.SelLength = Len(txtPlacaNumero.Text)
End Sub
Private Sub txtPlacaNumero_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtKM.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
    If Len(txtPlacaNumero.Text) = 4 Then
        txtKM.SetFocus
    End If
End Sub
