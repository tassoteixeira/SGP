VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form lst_quadro_funcionario 
   Caption         =   "Emissão do Quadro de Funcionários"
   ClientHeight    =   1935
   ClientLeft      =   3990
   ClientTop       =   2010
   ClientWidth     =   6795
   Icon            =   "lst_quadro_funcionario.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "lst_quadro_funcionario.frx":030A
   ScaleHeight     =   1935
   ScaleWidth      =   6795
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1140
      Picture         =   "lst_quadro_funcionario.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Visualiza quadro de funcionários."
      Top             =   960
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3000
      Picture         =   "lst_quadro_funcionario.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Imprime quadro de funcionários."
      Top             =   960
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4860
      Picture         =   "lst_quadro_funcionario.frx":3074
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   960
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6555
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2760
         Picture         =   "lst_quadro_funcionario.frx":4706
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin MSMask.MaskEdBox msk_data 
         Height          =   315
         Left            =   1680
         TabIndex        =   2
         Top             =   240
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label5 
         Caption         =   "Data de &Emissão"
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1515
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   1140
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "lst_quadro_funcionario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Início de variáveis padrão para relatório
Dim lLinha As Integer
Dim lPagina As Integer
Dim lLocal As Integer
Dim lNomeArquivo As String
'Fim de variáveis padrão para relatório
Dim tbl_funcionario As Table
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    tbl_funcionario.Close
End Sub
Private Sub ZeraVariaveis()
    lLinha = 0
    lPagina = 0
End Sub
Private Sub Relatorio()
    Dim x_linha As String
    Dim i As Integer
    ZeraVariaveis
    'Loop Funcionario
    With tbl_funcionario
        If .RecordCount > 0 Then
            If lPagina = 0 Then
                ImpCab
            End If
            For i = 0 To 5
                If i = 0 Then
                    BioImprime "@Printer.Print " & "|                            PERÍODO 0 (08:00 ÀS 18:00)                        |"
                ElseIf i = 1 Then
                    BioImprime "@Printer.Print " & "+------------------------------------------+----------------------+------------+"
                    BioImprime "@Printer.Print " & "|                            PERÍODO 1 (06:00 ÀS 14:00)                        |"
                ElseIf i = 2 Then
                    BioImprime "@Printer.Print " & "+------------------------------------------+----------------------+------------+"
                    BioImprime "@Printer.Print " & "|                            PERÍODO 2 (14:00 ÀS 22:00)                        |"
                ElseIf i = 3 Then
                    BioImprime "@Printer.Print " & "+------------------------------------------+----------------------+------------+"
                    BioImprime "@Printer.Print " & "|                            PERÍODO 3 (22:00 ÀS 06:00)                        |"
                ElseIf i = 4 Then
                    BioImprime "@Printer.Print " & "+------------------------------------------+----------------------+------------+"
                    BioImprime "@Printer.Print " & "|                            PERÍODO 4 (1/2/3)                                 |"
                ElseIf i = 5 Then
                    BioImprime "@Printer.Print " & "+------------------------------------------+----------------------+------------+"
                    BioImprime "@Printer.Print " & "|                            PERÍODO 5 (OUTROS)                                |"
                End If
                BioImprime "@Printer.Print " & "+------------------------------------------+----------------------+------------+"
                .Seek ">=", g_empresa, i, 0
                If Not .NoMatch Then
                    Do Until .EOF
                        If !Empresa <> g_empresa Or !Periodo <> i Then
                            Exit Do
                        End If
                        If !Situacao = "A" Then
                            x_linha = "|                                          |                      |            |"
                            Mid(x_linha, 3, 40) = !Nome
                            Mid(x_linha, 46, 20) = !Cargo
                            If Not IsNull(![Data de Admissao]) Then
                                Mid(x_linha, 69, 10) = Format(![Data de Admissao], "dd/mm/yyyy")
                            End If
                            BioImprime "@Printer.Print " & x_linha
                        End If
                        .MoveNext
                    Loop
                End If
            Next
            ImpTotal
            BioImprime "@@Printer.EndDoc"
            BioFechaImprime
            g_string = lLocal & lNomeArquivo & "|@|Relatório do Quadro de Funcionário|@|"
            frm_preview.Show 1
        End If
    End With
    cmd_sair.SetFocus
End Sub
Private Sub ImpTotal()
    Dim x_linha As String
    x_linha = "+------------------------------------------+----------------------+------------+"
    Mid(x_linha, 5, 22) = " Cerrado Informática. "
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & " "
End Sub
Private Sub ImpCab()
    Dim x_linha As String
    If lPagina = 0 Then
        lNomeArquivo = BioCriaImprime
        'seleciona medidas para centímetros
        BioImprime "@@Printer.ScaleMode = 7"
        BioImprime "@@Printer.PaperSize = 1"
        BioImprime "@@Printer.FontName = Draft 10cpi"
        BioImprime "@@Printer.FontName = Draft 10cpi"
        'teste para imprimir letra correta
        BioImprime "@@Printer.FontBold = False"
        BioImprime "@@ImprimeTexto " & Chr(34) & "  " & Chr(34) & ", 1, 2, 2, 1"
    End If
    lPagina = lPagina + 1
    lLinha = 0
    BioImprime "@@Printer.FontName = Draft 5cpi"
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@@Printer.CurrentY = 0"
    BioImprime "@Printer.Print " & "+------------------------------------------------------------------------------+"
    BioImprime "@@Printer.FontBold = True"
    x_linha = "|                                                                  Página, " & Format(lPagina, "000") & " |"
    Mid(x_linha, 3, 40) = g_nome_empresa
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontBold = False"
    BioImprime "@Printer.Print " & "| QUADRO DE FUNCIONARIOS                                   Goiânia, " & msk_data & " |"
    BioImprime "@Printer.Print " & "+------------------------------------------------------------------------------+"
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    msk_data = RetiraGString(1)
    cmd_visualizar.SetFocus
    g_string = " "
End Sub
Private Sub cmd_imprimir_Click()
    lLocal = 1
    If ValidaCampos Then
        If SelecionaImpressoraEpson(Me) Then
            Call GravaAuditoria(1, Me.name, 7, "")
            Relatorio
        End If
    End If
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If Not IsDate(msk_data) Then
        MsgBox "Informe a data de emissão.", 64, "Atenção!"
        msk_data.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub cmd_visualizar_Click()
    lLocal = 0
    If ValidaCampos Then
        If SelecionaImpressoraEpson(Me) Then
            Call GravaAuditoria(1, Me.name, 6, "")
            Relatorio
        End If
    End If
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    If Not IsDate(msk_data.Text) Then
        msk_data.Text = Format(g_data_def, "dd/mm/yyyy")
        cmd_visualizar.SetFocus
    End If
    Screen.MousePointer = 1
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        cmd_imprimir_Click
    ElseIf KeyCode = vbKeyF9 Then
        KeyCode = 0
        cmd_visualizar_Click
    End If
End Sub
Private Sub Form_Load()
    CentraForm Me
    Set tbl_funcionario = bd_sgp.OpenTable("Funcionario")
    tbl_funcionario.Index = "id_periodo"
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_visualizar.SetFocus
    End If
End Sub
