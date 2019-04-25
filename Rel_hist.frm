VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form lst_historico 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Relatório do Histórico"
   ClientHeight    =   1515
   ClientLeft      =   4005
   ClientTop       =   2055
   ClientWidth     =   3690
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
   Icon            =   "Rel_hist.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1515
   ScaleWidth      =   3690
   Begin VB.CommandButton cmd_data 
      Height          =   315
      Left            =   3060
      Picture         =   "Rel_hist.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Selecione a data pelo calendário."
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
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
      Left            =   600
      Picture         =   "Rel_hist.frx":15E4
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Imprime histórico."
      Top             =   540
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
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
      Left            =   2100
      Picture         =   "Rel_hist.frx":2BEE
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   540
      Width           =   795
   End
   Begin MSMask.MaskEdBox msk_data 
      Height          =   345
      Left            =   1860
      TabIndex        =   1
      Top             =   105
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
      _Version        =   393216
      BackColor       =   16777215
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Data do Relatório"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   1665
   End
End
Attribute VB_Name = "lst_historico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmd_data_Click()
    g_string = msk_data.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data.Text = RetiraGString(1)
        cmd_imprimir.SetFocus
    Else
        msk_data.Text = RetiraGString(1)
        cmd_imprimir.SetFocus
    End If
    g_string = " "
End Sub
Private Sub cmd_imprimir_Click()
    If ValidaCampos Then
        If SelecionaImpressoraHP(Me) Then
            Call GravaAuditoria(1, Me.name, 7, "")
            Relatorio
            cmd_sair.SetFocus
        End If
    End If
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If Not IsDate(msk_data) Then
        MsgBox "Informe a data do relatório.", 64, "Atenção!"
        msk_data.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    tbl_historico.Close
End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    msk_data.Text = Format(g_data_def, "dd/mm/yyyy")
    cmd_imprimir.SetFocus
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        cmd_imprimir_Click
    End If
End Sub
Private Sub Form_Load()
    Screen.MousePointer = 1
    centerform Me
    Set tbl_historico = bd_sgp.OpenTable("Historico")
End Sub
Private Sub Relatorio()
    Dim i As Integer
    Dim i2 As Integer
    Dim linha As String
    Dim posicao_y As Currency
    Dim posicao_x As Currency
    Dim x_pai As String
    Dim x_filho As String

    Printer.ScaleMode = 7
    Printer.FontName = "Arial"
    Printer.FontSize = 15
    Printer.Print
    Printer.Print
    linha = "Relatório do Histórico                                  Goiânia, " & msk_data
    Printer.Print linha
    Printer.Print
    linha = "Empresa: " & g_nome_empresa
    Printer.Print linha
    posicao_x = Printer.CurrentX
    posicao_y = Printer.CurrentY
    Printer.FontSize = 10

    'Load cadastro_historico
    For i = 0 To cadastro_historico!Outline1.ListCount - 1
        linha = ""
        For i2 = 0 To cadastro_historico!Outline1.Indent(i) * 3
            linha = linha & " "
        Next

        linha = linha & cadastro_historico!Outline1.List(i)
        Printer.CurrentY = posicao_y + 0.5
        posicao_y = Printer.CurrentY
        Printer.Print linha
        x_pai = cadastro_historico!Outline1.FullPath(i)
        x_filho = ""
        PaiFilho x_pai, x_filho
        tbl_historico.Index = "id_pai"
        tbl_historico.Seek "=", g_empresa, x_pai, x_filho
        If Not tbl_historico.NoMatch Then
            If tbl_historico![Debito Credito] = "C" Then
                linha = "Crédito"
            Else
                linha = "Débito"
            End If
        Else
            linha = ""
        End If
        Printer.CurrentY = posicao_y
        Printer.CurrentX = 15
        Printer.Print linha
    Next
    Unload cadastro_historico
    Printer.EndDoc
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_imprimir.SetFocus
    End If
End Sub
