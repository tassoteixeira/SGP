VERSION 5.00
Begin VB.Form calc_litro 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Calcula Quantidade de Litros"
   ClientHeight    =   2595
   ClientLeft      =   3330
   ClientTop       =   2790
   ClientWidth     =   3780
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
   Icon            =   "calc_litro.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Bancos"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2595
   ScaleWidth      =   3780
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
      Picture         =   "calc_litro.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Confirma o registro atual."
      Top             =   1620
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
      Picture         =   "calc_litro.frx":1914
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Cancela o registro atual."
      Top             =   1620
      Width           =   795
   End
   Begin VB.Frame frmDados
      Height          =   1395
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   3585
      ForeColor       =   0
      Begin VB.TextBox txt_total_venda 
         Height          =   285
         Left            =   2340
         MaxLength       =   10
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lbl_quantidade_litros 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2340
         TabIndex        =   5
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lbl_preco_venda 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2340
         TabIndex        =   1
         Top             =   180
         Width           =   1095
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Quantidade de litros"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1020
         Width           =   2175
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Preço de venda"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Total do valor da venda"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   660
         Width           =   2175
      End
   End
End
Attribute VB_Name = "calc_litro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmd_cancelar_Click()
    g_valor = 0
    g_string = ""
    Unload Me
End Sub
Private Sub cmd_ok_Click()
    g_valor = fValidaValor2(lbl_quantidade_litros.Caption)
    g_string = fValidaValor2(lbl_quantidade_litros.Caption) & "|@|"
    g_string = g_string & fValidaValor2(txt_total_venda) & "|@|"
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
    lbl_preco_venda.Caption = Format(g_valor, "###,##0.0000")
End Sub
Private Sub txt_total_venda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub txt_total_venda_LostFocus()
    txt_total_venda = Format(txt_total_venda, "###,##0.00")
    lbl_quantidade_litros.Caption = Format(fValidaValor2(txt_total_venda) / fValidaValor4(lbl_preco_venda.Caption), "###,##0.00")
End Sub
