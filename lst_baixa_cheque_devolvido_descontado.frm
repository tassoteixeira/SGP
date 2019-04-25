VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form lst_baixa_cheque_devolvido_descontado 
   Caption         =   "Emissão da Baixa dos Cheques Devolvidos Descontados"
   ClientHeight    =   2715
   ClientLeft      =   3990
   ClientTop       =   2010
   ClientWidth     =   6795
   Icon            =   "lst_baixa_cheque_devolvido_descontado.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "lst_baixa_cheque_devolvido_descontado.frx":030A
   ScaleHeight     =   2715
   ScaleWidth      =   6795
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1200
      Picture         =   "lst_baixa_cheque_devolvido_descontado.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Visualiza baixa de cheque devolvido descontado."
      Top             =   1740
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3000
      Picture         =   "lst_baixa_cheque_devolvido_descontado.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Imprime baixa de cheque devolvido descontado."
      Top             =   1740
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4800
      Picture         =   "lst_baixa_cheque_devolvido_descontado.frx":3074
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   1740
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   1515
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6555
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   5940
         Picture         =   "lst_baixa_cheque_devolvido_descontado.frx":4706
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2760
         Picture         =   "lst_baixa_cheque_devolvido_descontado.frx":59E0
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2760
         Picture         =   "lst_baixa_cheque_devolvido_descontado.frx":6CBA
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.ComboBox cbo_inativo 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1080
         Width           =   915
      End
      Begin MSMask.MaskEdBox msk_data_f 
         Height          =   315
         Left            =   4860
         TabIndex        =   8
         Top             =   660
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_data_i 
         Height          =   315
         Left            =   1680
         TabIndex        =   5
         Top             =   660
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
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
      Begin VB.Label Label3 
         Caption         =   "I&mprimir inativos"
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "&Data inicial"
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   660
         Width           =   1515
      End
      Begin VB.Label Label2 
         Caption         =   "D&ata final"
         Height          =   315
         Left            =   3960
         TabIndex        =   7
         Top             =   660
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Data de &emissão"
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1515
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   1980
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "lst_baixa_cheque_devolvido_descontado"
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
Dim tbl_baixa_cheque_devolvido_descontado As Table
Dim tbl_funcionario As Table
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    tbl_baixa_cheque_devolvido_descontado.Close
    tbl_funcionario.Close
End Sub
Private Sub PreencheCboInativo()
    cbo_inativo.Clear
    cbo_inativo.AddItem "Sim"
    cbo_inativo.ItemData(cbo_inativo.NewIndex) = 1
    cbo_inativo.AddItem "Não"
    cbo_inativo.ItemData(cbo_inativo.NewIndex) = 2
    cbo_inativo.AddItem "Geral"
    cbo_inativo.ItemData(cbo_inativo.NewIndex) = 3
End Sub
Private Sub PreencheDadosInicial()
    msk_data = Format(g_data_def, "dd/mm/yyyy")
    msk_data_i = Format(g_data_def - 1, "dd/mm/yyyy")
    msk_data_f = Format(g_data_def - 1, "dd/mm/yyyy")
    cbo_inativo.ListIndex = 2
    With tbl_baixa_cheque_devolvido_descontado
        If .RecordCount > 0 Then
            .Seek ">=", g_empresa, CDate("01/01/1900"), 0, "    ", "      "
            If Not .NoMatch Then
                If !Empresa = g_empresa Then
                    msk_data_i = Format(![Data do Pagamento], "dd/mm/yyyy")
                End If
            End If
            .Seek "<=", g_empresa, CDate("31/12/2500"), 9999, "9999", "999999"
            If Not .NoMatch Then
                If !Empresa = g_empresa Then
                    msk_data_f = Format(![Data do Pagamento], "dd/mm/yyyy")
                End If
            End If
        End If
    End With
End Sub
Private Sub ZeraVariaveis()
    lLinha = 0
    lPagina = 0
    lTotal = 0
    lTotalQtd = 0
End Sub
Private Sub Relatorio()
    ZeraVariaveis
    'Verifica movimento
    With tbl_baixa_cheque_devolvido_descontado
        If .RecordCount > 0 Then
            .Seek ">=", g_empresa, CDate(msk_data_i), 0, "    ", "      "
            If Not .NoMatch Then
                If !Empresa = g_empresa And ![Data do Pagamento] <= CDate(msk_data_f) Then
                    ImpDados
                End If
            End If
        End If
    End With
    cmd_sair.SetFocus
End Sub
Private Sub ImpDados()
    Dim x_linha As String * 137
    'loop baixa de cheques devolvido descontado
    With tbl_baixa_cheque_devolvido_descontado
        Do Until .EOF
            If !Empresa <> g_empresa Or ![Data do Pagamento] > CDate(msk_data_f) Then
                Exit Do
            End If
            If (cbo_inativo = "Sim" And !Inativo = True) Or (cbo_inativo = "Não" And !Inativo = False) Or cbo_inativo = "Geral" Then
                If lPagina = 0 Then
                    ImpCab
                End If
                If lLinha >= 57 Then
                    x_linha = "+-----+------+------+------------+--------------------------------------+---+----------------------+----------+----------+------------+-+"
                    Mid(x_linha, 40, 22) = " Cerrado Informática. "
                    BioImprime "@Printer.Print " & x_linha
                    BioImprime "@@Printer.NewPage"
                    ImpCab
                End If
                ImpDet
                lTotal = lTotal + !valor
                lTotalQtd = lTotalQtd + 1
            End If
            .MoveNext
        Loop
    End With
    If lTotal > 0 Then
        ImpTotal
        BioImprime "@@Printer.EndDoc"
        BioFechaImprime
        g_string = lLocal & lNomeArquivo & "|@|Relatório dos Cheques Devolvido Descontado|@|"
        frm_preview.Show 1
    End If
End Sub
Private Sub ImpDet()
    Dim x_nome_funcionario As String * 40
    Dim x_linha As String * 137
    Dim i As Integer
    x_linha = "|     |      |      |            |                                      |   |                      |          |          |            | |"
    With tbl_baixa_cheque_devolvido_descontado
        i = Len(Format(![Numero do Banco], "###"))
        Mid(x_linha, 3 + 3 - i, i) = Format(![Numero do Banco], "###")
        i = Len(Format(![Numero da Agencia], "####"))
        Mid(x_linha, 9 + 4 - i, i) = Format(![Numero da Agencia], "####")
        i = Len(Format(![Numero do Cheque], "######"))
        Mid(x_linha, 15 + 6 - i, i) = Format(![Numero do Cheque], "######")
        i = Len(Format(!valor, "###,##0.00"))
        Mid(x_linha, 23 + 10 - i, i) = Format(!valor, "###,##0.00")
        Mid(x_linha, 35, 38) = !Emitente
        i = Len(Format(![Codigo do Funcionario], "###"))
        Mid(x_linha, 74 + 3 - i, i) = Format(![Codigo do Funcionario], "###")
        x_nome_funcionario = Space(40)
        tbl_funcionario.Seek "=", g_empresa, ![Codigo do Funcionario]
        If Not tbl_funcionario.NoMatch Then
            x_nome_funcionario = tbl_funcionario!Nome
        Else
            x_nome_funcionario = "** Não Cadastrado **"
        End If
        Mid(x_linha, 78, 22) = x_nome_funcionario
        Mid(x_linha, 101, 10) = Format(![Data da Entrega], "dd/mm/yyyy")
        Mid(x_linha, 112, 10) = Format(![Data do Pagamento], "dd/mm/yyyy")
        Mid(x_linha, 123, 12) = !Motivo
        If !Inativo Then
            Mid(x_linha, 136, 1) = "I"
        Else
            Mid(x_linha, 136, 1) = "A"
        End If
    End With
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@Printer.Print " & x_linha
    lLinha = lLinha + 1
End Sub
Private Sub ImpTotal()
    Dim x_linha As String * 137
    Dim i As Integer
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@Printer.Print " & "+-----+------+------+------------+--------------------------------------+---+----------------------+----------+----------+------------+-+"
    x_linha = "|         *** TOTAL |            |                                                                                                      |"
    i = Len(Format(lTotal, "###,##0.00"))
    Mid(x_linha, 23 + 10 - i, i) = Format(lTotal, "###,##0.00")
    BioImprime "@@Printer.FontBold = True"
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontBold = False"
    x_linha = "+-------------------+------------+------------------------------------------------------------------------------------------------------+"
    Mid(x_linha, 40, 22) = " Cerrado Informática. "
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & "  "
End Sub
Private Sub ImpCab()
    Dim x_linha As String * 137
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
    x_linha = "|                                                                  Página,     |"
    x_string_40 = g_nome_empresa
    Mid(x_linha, 3, 40) = x_string_40
    Mid(x_linha, 76, 3) = Format(lPagina, "000")
    BioImprime "@@Printer.FontBold = True"
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontBold = False"
    x_linha = "| RELAÇÃO DA BAIXA DOS CHEQUES DEVOLVIDOS DESCONTADOS      Goiânia,            |"
    Mid(x_linha, 69, 10) = msk_data
    BioImprime "@Printer.Print " & x_linha
    x_linha = "| Referente a.:            a                                                   |"
    Mid(x_linha, 17, 10) = msk_data_i
    Mid(x_linha, 30, 10) = msk_data_f
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@@Printer.FontBold = False"
    BioImprime "@Printer.Print " & "+-----+------+------+------------+--------------------------------------+---+----------------------+----------+----------+------------+-+"
    BioImprime "@Printer.Print " & "|BANCO| AGEN.|N. CH.| VALOR  CH. |NOME DO EMITENTE                      |COD| NOME DO FUNCIONÁRIO  |  ENTREGA | PAGAMENTO|MOTIVO      |*|"
    BioImprime "@Printer.Print " & "+-----+------+------+------------+--------------------------------------+---+----------------------+----------+----------+------------+-+"
End Sub
Private Sub cbo_inativo_GotFocus()
    SendMessageLong cbo_inativo.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_inativo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_imprimir.SetFocus
    End If
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i = RetiraGString(1)
        msk_data_f = RetiraGString(2)
        cbo_inativo.SetFocus
    Else
        msk_data = RetiraGString(1)
        msk_data_i.SetFocus
    End If
    g_string = " "
End Sub
Private Sub cmd_data_f_Click()
    g_string = msk_data_f
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i = RetiraGString(1)
        msk_data_f = RetiraGString(2)
    Else
        msk_data_f = RetiraGString(1)
    End If
    g_string = " "
    cbo_inativo.SetFocus
End Sub
Private Sub cmd_data_i_Click()
    g_string = msk_data_i
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i = RetiraGString(1)
        msk_data_f = RetiraGString(2)
        cbo_inativo.SetFocus
    Else
        msk_data_i = RetiraGString(1)
        msk_data_f.SetFocus
    End If
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
    ElseIf Not IsDate(msk_data_i) Then
        MsgBox "Informe a data inicial.", 64, "Atenção!"
        msk_data_i.SetFocus
    ElseIf Not IsDate(msk_data_f) Then
        MsgBox "Informe a data final.", 64, "Atenção!"
        msk_data_f.SetFocus
    ElseIf CDate(msk_data_f) < CDate(msk_data_i) Then
        MsgBox "Data final deve ser maior que a data inicial.", 64, "Atenção!"
        msk_data_f.SetFocus
    ElseIf cbo_inativo.ListIndex = -1 Then
        MsgBox "Escolha o tipo de impressão.", 64, "Atenção!"
        cbo_inativo.SetFocus
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
    If Not IsDate(msk_data) Then
        PreencheDadosInicial
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
    Set tbl_baixa_cheque_devolvido_descontado = bd_sgp.OpenTable("Baixa_Cheque_Devolvido_Descontado")
    Set tbl_funcionario = bd_sgp.OpenTable("Funcionario")
    tbl_baixa_cheque_devolvido_descontado.Index = "id_data_pagamento"
    tbl_funcionario.Index = "id_codigo"
    PreencheCboInativo
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_f_GotFocus()
    msk_data_f.SelStart = 0
    msk_data_f.SelLength = 2
End Sub
Private Sub msk_data_f_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_inativo.SetFocus
    End If
End Sub
Private Sub msk_data_i_GotFocus()
    msk_data_i.SelStart = 0
    msk_data_i.SelLength = 2
End Sub
Private Sub msk_data_i_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data_f.SetFocus
    End If
End Sub
Private Sub msk_data_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data_i.SetFocus
    End If
End Sub
