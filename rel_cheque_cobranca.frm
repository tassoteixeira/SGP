VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form relatorio_cheque_cobranca 
   Caption         =   "Relação dos Cheques em Cobrança"
   ClientHeight    =   1875
   ClientLeft      =   1875
   ClientTop       =   1725
   ClientWidth     =   4515
   Icon            =   "rel_cheque_cobranca.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "rel_cheque_cobranca.frx":030A
   ScaleHeight     =   1875
   ScaleWidth      =   4515
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3120
      Picture         =   "rel_cheque_cobranca.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   900
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   1860
      Picture         =   "rel_cheque_cobranca.frx":19E2
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Imprime cheques em cobrança."
      Top             =   900
      Width           =   795
   End
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   600
      Picture         =   "rel_cheque_cobranca.frx":2FEC
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Visualiza cheques em cobrança."
      Top             =   900
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   4275
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   3180
         Picture         =   "rel_cheque_cobranca.frx":4706
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin MSMask.MaskEdBox msk_data 
         Height          =   315
         Left            =   2100
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
         Caption         =   "Data de Emissão"
         Height          =   315
         Left            =   540
         TabIndex        =   1
         Top             =   240
         Width           =   1515
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   900
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "relatorio_cheque_cobranca"
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
Dim lTotal As Currency
Dim lTotalQtd As Currency
Dim tbl_cheque_cobranca As Table
Dim l_sql As String
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    tbl_cheque_cobranca.Close
End Sub
Private Sub Relatorio()
    ZeraVariaveis
    'Verifica movimento
    tbl_cheque_cobranca.Seek ">", g_empresa, ""
    If Not tbl_cheque_cobranca.NoMatch Then
        If tbl_cheque_cobranca!Empresa = g_empresa Then
            ImpDados
        End If
    End If
    cmd_sair.SetFocus
End Sub
Private Sub ImpDados()
    Dim x_linha As String * 137
    'loop movimento de cheques
    With tbl_cheque_cobranca
        Do Until .EOF
            If !Empresa <> g_empresa Then
                Exit Do
            End If
            If lPagina = 0 Then
                ImpCab
            End If
            If lLinha >= 57 Then
                x_linha = "+-------+-------+------------+--------+---------------+------------------------------------------+----------------------+---------------+"
                Mid(x_linha, 59, 22) = " Cerrado Informatica. "
                BioImprime "@Printer.Print " & x_linha
                BioImprime "@@Printer.NewPage"
                ImpCab
            End If
            ImpDet
            lTotal = lTotal + !valor
            lTotalQtd = lTotalQtd + 1
            .MoveNext
        Loop
    End With
    If lTotal > 0 Then
        ImpTotal
        BioImprime "@@Printer.EndDoc"
        BioFechaImprime
        g_string = lLocal & lNomeArquivo & "|@|Relatório de Cheque em Cobrança|@|"
        frm_preview.Show 1
    End If
End Sub
Private Sub ImpDet()
    Dim x_linha As String * 137
    Dim i As Integer
    x_linha = "|       |       |            |        |               |                                          |                      |               |"
    With tbl_cheque_cobranca
        i = Len(Format(![Numero do Banco], "###"))
        Mid(x_linha, 4 + 3 - i, i) = Format(![Numero do Banco], "###")
        i = Len(Format(![Numero da Agencia], "####"))
        Mid(x_linha, 12 + 4 - i, i) = Format(![Numero da Agencia], "####")
        Mid(x_linha, 19, 10) = ![Numero da Conta]
        i = Len(Format(![Numero do Cheque], "######"))
        Mid(x_linha, 32 + 6 - i, i) = Format(![Numero do Cheque], "######")
        i = Len(Format(!valor, "##,###,##0.00"))
        Mid(x_linha, 41 + 13 - i, i) = Format(!valor, "##,###,##0.00")
        Mid(x_linha, 57, 40) = !Emitente
        Mid(x_linha, 100, 20) = !Motivo
    End With
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@Printer.Print " & x_linha
    lLinha = lLinha + 1
End Sub
Private Sub ImpCab()
    Dim x_string_40 As String * 40
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
    x_string_40 = g_nome_empresa
    BioImprime "@Printer.Print " & "| " & x_string_40 & "                         Página, " & Format(lPagina, "000") & " |"
    BioImprime "@Printer.Print " & "| RELAÇÃO DOS CHEQUES EM COBRANÇA                          Goiânia, " & msk_data & " |"
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@@Printer.FontBold = False"
    BioImprime "@Printer.Print " & "+-------+-------+------------+--------+---------------+------------------------------------------+----------------------+---------------+"
    BioImprime "@Printer.Print " & "| BANCO |AGENCIA| N.DA CONTA | CHEQUE |VALOR DO CHEQUE| NOME DO EMITENTE                         | MOTIVO               |               |"
    BioImprime "@Printer.Print " & "+-------+-------+------------+--------+---------------+------------------------------------------+----------------------+---------------+"
End Sub
Private Sub ImpTotal()
    Dim x_linha As String * 137
    Dim i As Integer
    BioImprime "@Printer.Print " & "+-------+-------+------------+--------+---------------+------------------------------------------+----------------------+---------------+"
    x_linha = "|                            ** TOTAL |               |                                          |                                      |"
    i = Len(Format(lTotal, "##,###,##0.00"))
    Mid(x_linha, 41 + 13 - i, i) = Format(lTotal, "##,###,##0.00")
    Mid(x_linha, 57, 20) = "** NÚMERO DE CHEQUES"
    i = Len(Format(lTotalQtd, "####"))
    Mid(x_linha, 78 + 4 - i, i) = Format(lTotalQtd, "####")
    BioImprime "@Printer.Print " & x_linha
    x_linha = "+-------------------------------------+---------------+------------------------------------------+--------------------------------------+"
    Mid(x_linha, 59, 22) = " Cerrado Informática. "
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & "  "
End Sub
Private Sub ZeraVariaveis()
    lLinha = 0
    lPagina = 0
    lTotal = 0
    lTotalQtd = 0
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    msk_data = RetiraGString(1)
    cmd_imprimir.SetFocus
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
        cmd_imprimir.SetFocus
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
    Set tbl_cheque_cobranca = bd_sgp.OpenTable("Cheque_Cobranca")
    tbl_cheque_cobranca.Index = "id_emitente"
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
