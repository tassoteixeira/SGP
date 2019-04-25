VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form emissao_memoria_fiscal 
   Caption         =   "Emissão da Memória Fiscal"
   ClientHeight    =   2295
   ClientLeft      =   3990
   ClientTop       =   2010
   ClientWidth     =   6795
   Icon            =   "emissao_memoria_fiscal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   Picture         =   "emissao_memoria_fiscal.frx":030A
   ScaleHeight     =   2295
   ScaleWidth      =   6795
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   1740
      Picture         =   "emissao_memoria_fiscal.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Imprime a Leitura da Memória Fiscal."
      Top             =   1320
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4260
      Picture         =   "emissao_memoria_fiscal.frx":195A
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   1320
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6555
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   5940
         Picture         =   "emissao_memoria_fiscal.frx":2FEC
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2760
         Picture         =   "emissao_memoria_fiscal.frx":42C6
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2760
         Picture         =   "emissao_memoria_fiscal.frx":55A0
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
      Begin VB.Label Label2 
         Caption         =   "D&ata final"
         Height          =   255
         Left            =   3840
         TabIndex        =   7
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "&Data inicial"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1515
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
      Top             =   1740
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "emissao_memoria_fiscal"
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
Dim lImpBematech As Boolean
Dim lImpSchalter As Boolean
Dim lImpMecaf As Boolean
Dim lImpQuick As Boolean
Dim lImpElgin As Boolean
Dim BemaRetorno As Integer
'Fim de variáveis padrão para relatório
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    If lImpMecaf Then
        CloseCif
    End If
End Sub
Function TestaImpressora() As Boolean
    Dim dados As String
    'Dim NumeroArquivo As Integer
    'NumeroArquivo = FreeFile
    TestaImpressora = False
    'On Error GoTo FileError
    'Open "C:\VB5\SGP\CUPOM_DEMONSTRACAO.TXT" For Input As NumeroArquivo
    'Line Input #NumeroArquivo, dados
    'lImpBematech = True
    'lImpSchalter = False
    'lImpMecaf = False
    'If Not EOF(NumeroArquivo) Then
    '    Line Input #NumeroArquivo, dados
    '    If Mid(dados, 1, 5) = "MECAF" Then
    '        lImpBematech = False
    '        lImpSchalter = False
    '        lImpMecaf = True
    '    ElseIf Mid(dados, 1, 8) = "SCHALTER" Then
    '        lImpBematech = False
    '        lImpMecaf = False
    '        lImpSchalter = True
    '    End If
    'End If
    'Close #NumeroArquivo
    
    lImpBematech = False
    lImpSchalter = False
    lImpMecaf = False
    dados = ReadINI("CUPOM FISCAL", "Impressora Fiscal", gArquivoIni)
    Me.Caption = Me.Caption & " - ECF: " & dados
    If dados = "BEMATECH" Then
        lImpBematech = True
    ElseIf dados = "SCHALTER" Then
        lImpSchalter = True
    ElseIf dados = "MECAF" Then
        lImpMecaf = True
    ElseIf dados = "QUICK" Then
        lImpQuick = True
    ElseIf dados = "ELGIN" Then
        lImpElgin = True
    End If
    
    Exit Function
FileError:
    Exit Function
End Function
Private Sub ZeraVariaveis()
    lLinha = 0
    lPagina = 0
End Sub
Private Sub Relatorio()
    Dim res As Byte
    Dim Retorno As Long
    Dim x_data_i As String
    Dim x_data_f As String
    Dim x_numero_i As String
    Dim x_numero_f As String
    Dim x_tipo_leitura As Integer
    Dim x_nome_arquivo As String
    Dim x_retorno As Integer
    
    ZeraVariaveis
    If lImpBematech Then
        BemaRetorno = Bematech_FI_LeituraMemoriaFiscalData(Format(msk_data_i.Text, "dd/mm/yyyy"), Format(msk_data_f.Text, "dd/mm/yyyy"))
        'Call Abre_ProtocoloCF(1)
        'ComandoCF = Chr(27) + "|08|" + Format(CDate(msk_data_i), "dd") + Format(CDate(msk_data_i), "mm") + Format(CDate(msk_data_i), "yy") + "|" + Format(CDate(msk_data_f), "dd") + Format(CDate(msk_data_f), "mm") + Format(CDate(msk_data_f), "yy") + "|" + "I" + "|" + Chr(27)
        'Envia_ComandoCF
    ElseIf lImpMecaf Then
        Retorno = OpenCif
        If Retorno <> 0 Then
            If Retorno = -92 Then
                MsgBox "A Impressora está acusando falta de papel!" & Chr(10) & "Favor abrir a tampa trazeira e verificar.", vbInformation, "Falta de Papel!"
                Exit Sub
            End If
        End If
        If MsgBox("Leitura resumida?", vbQuestion + vbYesNo + vbDefaultButton1, "Tipo de Leitura!") = vbYes Then
            res = Asc("1")
        Else
            res = Asc("0")
        End If
        x_data_i = Format(msk_data_i.Text, "ddmmyy")
        x_data_f = Format(msk_data_f.Text, "ddmmyy")
        MsgBox "Res:" & res & " Data i:" & x_data_i & " Data f:" & x_data_f
        Retorno = LeMemFiscalData(x_data_i, x_data_f, res)
        Sleep 25000
    ElseIf lImpSchalter Then
        x_tipo_leitura = 1
        x_data_i = Format(msk_data_i.Text, "ddmmyy")
        x_data_f = Format(msk_data_f.Text, "ddmmyy")
        x_numero_i = "0000"
        x_numero_f = "0000"
        x_nome_arquivo = "teste.txt"
        x_retorno = ecfLeitMemFisc(x_tipo_leitura, x_data_i, x_data_f, x_numero_i, x_numero_f, x_nome_arquivo)
    ElseIf lImpQuick Then
        If EcfQuickEmiteMemoriaFiscal(CDate(msk_data_i.Text), CDate(msk_data_f.Text), "I", False) Then
            BemaRetorno = 1
        Else
            BemaRetorno = 0
        End If
    ElseIf lImpElgin Then
        BemaRetorno = Elgin_LeituraMemoriaFiscalData(msk_data_i.Text, msk_data_f.Text, "c")
    End If
    cmd_sair.SetFocus
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    msk_data.Text = RetiraGString(1)
    cmd_imprimir.SetFocus
    g_string = ""
End Sub
Private Sub cmd_data_f_Click()
    g_string = msk_data_f.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i.Text = RetiraGString(1)
        msk_data_f.Text = RetiraGString(2)
    Else
        msk_data_f.Text = RetiraGString(1)
    End If
    g_string = ""
    cmd_imprimir.SetFocus
End Sub
Private Sub cmd_data_i_Click()
    g_string = msk_data_i.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i.Text = RetiraGString(1)
        msk_data_f.Text = RetiraGString(2)
        cmd_imprimir.SetFocus
    Else
        msk_data_i.Text = RetiraGString(1)
        msk_data_f.SetFocus
    End If
    g_string = ""
End Sub
Private Sub cmd_imprimir_Click()
    lLocal = 1
    If ValidaCampos Then
        'If SelecionaImpressoraEpson(Me) Then
            Call GravaAuditoria(1, Me.name, 7, "")
            Relatorio
        'End If
    End If
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If Not IsDate(msk_data.Text) Then
        MsgBox "Informe a data de emissão.", vbInformation, "Atenção!"
        msk_data.SetFocus
    ElseIf Not IsDate(msk_data_i.Text) Then
        MsgBox "Informe a data inicial.", vbInformation, "Atenção!"
        msk_data_i.SetFocus
    ElseIf Not IsDate(msk_data_f.Text) Then
        MsgBox "Informe a data final.", vbInformation, "Atenção!"
        msk_data_f.SetFocus
    ElseIf CDate(msk_data_f.Text) < CDate(msk_data_i.Text) Then
        MsgBox "Data final deve ser maior ou igual a " & CDate(msk_data_i.Text) & ".", vbInformation, "Atenção!"
        msk_data_f.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    If Not IsDate(msk_data.Text) Then
        msk_data.Text = Format(g_data_def, "dd/mm/yyyy")
        msk_data_i.Text = Format(g_data_def, "dd/mm/yyyy")
        msk_data_f.Text = Format(g_data_def, "dd/mm/yyyy")
        cmd_imprimir.SetFocus
    End If
    Screen.MousePointer = 1
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        cmd_imprimir_Click
    End If
End Sub
Private Sub Form_Load()
    CentraForm Me
    TestaImpressora
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
        cmd_imprimir.SetFocus
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
