VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form emissao_plano_conta 
   Caption         =   "Emissão do Plano de Contas"
   ClientHeight    =   2250
   ClientLeft      =   3990
   ClientTop       =   2010
   ClientWidth     =   5715
   Icon            =   "emissao_plano_conta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "emissao_plano_conta.frx":030A
   ScaleHeight     =   2250
   ScaleWidth      =   5715
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1073
      Picture         =   "emissao_plano_conta.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Visualiza plano de contas."
      Top             =   1320
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   2453
      Picture         =   "emissao_plano_conta.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Imprime plano de contas."
      Top             =   1320
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3833
      Picture         =   "emissao_plano_conta.frx":3074
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   1320
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   675
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5475
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2760
         Picture         =   "emissao_plano_conta.frx":4706
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
         Caption         =   "&Data de Emissão"
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1515
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   1740
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "emissao_plano_conta"
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
Dim lCodigoGrupo As String
Dim lCodigoAnterior As String
Dim lSQl As String
Private rsTabela As New adodb.Recordset
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
End Sub
Private Sub ZeraVariaveis()
    lLinha = 0
    lPagina = 0
    lCodigoGrupo = ""
    lCodigoAnterior = ""
End Sub
Private Sub Relatorio()
    ZeraVariaveis
    'Prepara SQL
    lSQl = ""
    lSQl = lSQl & "SELECT Codigo, Nome, [Conta Reduzida]"
    lSQl = lSQl & "  FROM Plano_Conta"
    lSQl = lSQl & " WHERE Empresa = " & g_empresa
    lSQl = lSQl & " ORDER BY [Codigo do Grupo] ASC, [Tipo da Conta] DESC, Nome ASC, Codigo ASC"
    'Abre RecordSet
    Set rsTabela = New adodb.Recordset
    Set rsTabela = Conectar.RsConexao(lSQl)
    'Verifica movimento
    If rsTabela.RecordCount > 0 Then
        ImpDados
    End If
    If rsTabela.State = 1 Then
        rsTabela.Close
    End If
    cmd_sair.SetFocus
End Sub
Private Sub ImpDados()
    LoopTabela
    If lPagina > 0 Then
        ImpTotal
        BioImprime "@@Printer.EndDoc"
        BioFechaImprime
        g_string = lLocal & lNomeArquivo & "|@|Relatório do Plano de Contas|@|"
        frm_preview.Show 1
    End If
End Sub
Private Sub LoopTabela()
    'loop tabela
    Do Until rsTabela.EOF
        If lPagina = 0 Then
            ImpCab
        End If
        Call ImpDet
        rsTabela.MoveNext
    Loop
    ImpContaNova
End Sub
Private Sub ImpDet()
    Dim xLinha As String
    Dim i As Integer
    xLinha = "|               |         |                                                    |"
    If lLinha > 0 And lCodigoGrupo <> Mid(rsTabela("Codigo").Value, 1, 5) Then
        ImpContaNova
        BioImprime "@Printer.Print " & xLinha
        lLinha = lLinha + 1
    End If
    
    If lLinha >= 60 Then
         xLinha = "+---------------+---------+----------------------------------------------------+"
        Mid(xLinha, 31, 22) = " Cerrado Informática. "
        BioImprime "@Printer.Print " & xLinha
        BioImprime "@@Printer.NewPage"
        ImpCab
    End If
    
    lCodigoGrupo = Mid(rsTabela("Codigo").Value, 1, 5)
    xLinha = "|               |         |                                                    |"
    Mid(xLinha, 3, 13) = fMascaraContaContabil(rsTabela("Codigo").Value)
    If rsTabela("Conta Reduzida").Value > 0 Then
        i = Len(Format(rsTabela("Conta Reduzida").Value, "##,##0"))
        Mid(xLinha, 19 + 6 - i, i) = Format(rsTabela("Conta Reduzida").Value, "##,##0")
    End If
    Mid(xLinha, 29, 40) = rsTabela("Nome").Value
    'BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 1
    If rsTabela("Codigo").Value > lCodigoAnterior Then
        lCodigoAnterior = rsTabela("Codigo").Value
    End If
End Sub
Private Sub ImpContaNova()
    Dim xLinha As String
    xLinha = "|               |         |                                                    |"
    If Len(lCodigoAnterior) = 9 Then
        Mid(xLinha, 3, 13) = fMascaraContaContabil(CStr(CLng(lCodigoAnterior) + 1))
        Mid(xLinha, 29, 40) = String(40, "_")
        BioImprime "@Printer.Print " & xLinha
        lLinha = lLinha + 1
    End If
End Sub
Private Sub ImpTotal()
'    Dim y_local As Single
    Dim xLinha As String
    Dim i As Integer
    xLinha = "+---------------+--------------------------------------------------------------+"
    Mid(xLinha, 21, 22) = " Cerrado Informática. "
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & " "
End Sub
Private Sub ImpCab()
    Dim xLinha As String
    Dim i As Integer
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
    BioImprime "@@Printer.Print " & Chr(34) & " " & Chr(34)
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@@Printer.CurrentY = 0"
    xLinha = "+------------------------------------------------------------------------------+"
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontBold = True"
    xLinha = "|                                                                  Página: ___ |"
    Mid(xLinha, 3, 40) = g_nome_empresa
    Mid(xLinha, 76, 3) = Format(lPagina, "000")
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontBold = False"
    '                  1         2         3         4         5         6         7         8         9        10        11        12        13     13
    '         12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567
    '                                                                                                            123456789012345678901234567890
    xLinha = "| PLANO DE CONTAS                                           CIDADE, __/__/____ |"
    'Mid(xLinha, 68, 30) = cbo_grupo
    i = Len(g_cidade_empresa)
    Mid(xLinha, 37 + 30 - i, i) = g_cidade_empresa
    Mid(xLinha, 69, 10) = msk_data.Text
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontName = Draft 10cpi"
    xLinha = "+---------------+---------+----------------------------------------------------+"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| CÓDIGO  CONTA |  REDUZ. | DISCRIMINAÇÃO DA CONTA                             |"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "+---------------+---------+----------------------------------------------------+"
    BioImprime "@Printer.Print " & xLinha
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
    If Not IsDate(msk_data.Text) Then
        MsgBox "Informe a data de emissão.", vbInformation, "Atenção!"
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
