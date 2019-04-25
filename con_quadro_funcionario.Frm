VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form consulta_quadro_funcionario 
   Caption         =   "Consulta Quadro de Funcionários"
   ClientHeight    =   6570
   ClientLeft      =   2220
   ClientTop       =   2370
   ClientWidth     =   8655
   Icon            =   "con_quadro_funcionario.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "con_quadro_funcionario.frx":0442
   ScaleHeight     =   6570
   ScaleWidth      =   8655
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   7860
      Picture         =   "con_quadro_funcionario.frx":0488
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   5700
      Width           =   735
   End
   Begin VB.Frame frmDados
      Height          =   6495
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   7695
      Begin RichTextLib.RichTextBox txt_preview 
         Height          =   6315
         Left            =   60
         TabIndex        =   1
         Top             =   120
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   11139
         _Version        =   393217
         BackColor       =   16777215
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"con_quadro_funcionario.frx":1B1A
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
   End
End
Attribute VB_Name = "consulta_quadro_funcionario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flag_consulta_quadro_funcionario As Integer
Dim tbl_funcionario As Table
Private Sub Finaliza()
    tbl_funcionario.Close
End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub Form_Activate()
    If flag_consulta_quadro_funcionario = 0 Then
        Inicializa
        cmd_sair.SetFocus
    Else
        flag_consulta_quadro_funcionario = 0
    End If
End Sub
Private Sub Inicializa()
    Dim i As Integer
    Dim x_string As String
    txt_preview.Text = ""
    txt_preview.Text = txt_preview.Text & "+-----------------------------------------------------------------+" & Chr(13) & Chr(10)
    With tbl_funcionario
        If .RecordCount > 0 Then
            For i = 0 To 5
                If i = 0 Then
                    txt_preview.Text = txt_preview.Text & "|                            PERÍODO 0 (INTEGRAL)                 |" & Chr(13) & Chr(10)
                ElseIf i = 1 Then
                    txt_preview.Text = txt_preview.Text & "|                            PERÍODO 1 (06:00 ÀS 14:00)           |" & Chr(13) & Chr(10)
                ElseIf i = 2 Then
                    txt_preview.Text = txt_preview.Text & "|                            PERÍODO 2 (14:00 ÀS 22:00)           |" & Chr(13) & Chr(10)
                ElseIf i = 3 Then
                    txt_preview.Text = txt_preview.Text & "|                            PERÍODO 3 (22:00 ÀS 06:00)           |" & Chr(13) & Chr(10)
                ElseIf i = 4 Then
                    txt_preview.Text = txt_preview.Text & "|                            PERÍODO 4 (1/2/3)                    |" & Chr(13) & Chr(10)
                ElseIf i = 5 Then
                    txt_preview.Text = txt_preview.Text & "|                            PERÍODO 5 (OUTROS)                   |" & Chr(13) & Chr(10)
                End If
                txt_preview.Text = txt_preview.Text & "+------------------------------------------+----------------------+" & Chr(13) & Chr(10)
                .Seek ">=", g_empresa, i, 0
                If Not .NoMatch Then
                    Do Until .EOF
                        If !Empresa <> g_empresa Or !Periodo <> i Then
                            Exit Do
                        End If
                        If !Situacao = "A" Then
                            x_string = Space(80)
                            Mid(x_string, 1, 1) = "|"
                            Mid(x_string, 3, 40) = !Nome
                            Mid(x_string, 44, 1) = "|"
                            Mid(x_string, 46, 20) = !Cargo
                            Mid(x_string, 67, 1) = "|"
                            txt_preview.Text = txt_preview.Text & x_string & Chr(13) & Chr(10)
                        End If
                        .MoveNext
                    Loop
                End If
                txt_preview.Text = txt_preview.Text & "+------------------------------------------+----------------------+" & Chr(13) & Chr(10)
            Next
        End If
    End With
End Sub
Private Sub Form_Deactivate()
    flag_consulta_quadro_funcionario = 1
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        cmd_sair_Click
    End If
End Sub
Private Sub Form_Load()
    Screen.MousePointer = 1
    CentraForm Me
    Set tbl_funcionario = bd_sgp.OpenTable("Funcionario")
    tbl_funcionario.Index = "id_periodo"
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
