VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form analizador_tef 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transferência Eletrônica de Fundos - Cerrado Informática."
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8235
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "analizador_tef.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   8235
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frm_cnc 
      Caption         =   "Dados do COMPROVANTE para Cancelamento de Vendas TEF"
      Height          =   1560
      Left            =   1200
      TabIndex        =   2
      Top             =   900
      Visible         =   0   'False
      Width           =   5895
      Begin VB.CommandButton cmd_ok 
         Caption         =   "&OK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2820
         TabIndex        =   7
         Top             =   1080
         Width           =   795
      End
      Begin VB.TextBox txt_nsu 
         Height          =   300
         Left            =   2820
         MaxLength       =   10
         TabIndex        =   6
         ToolTipText     =   "Informe o número da transação (NSU -Número Sequencial Único)."
         Top             =   660
         Width           =   855
      End
      Begin VB.TextBox txt_valor 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   2820
         MaxLength       =   10
         TabIndex        =   4
         ToolTipText     =   "Informe o valor do cupom a ser cancelado."
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Número da Transação (NSU)"
         Height          =   300
         Left            =   180
         TabIndex        =   5
         Top             =   660
         Width           =   2595
      End
      Begin VB.Label Label2 
         Caption         =   "Valor do Cupom a ser Cancelado"
         Height          =   300
         Left            =   180
         TabIndex        =   3
         Top             =   300
         Width           =   2595
      End
   End
   Begin VB.Frame frm_procedimento 
      Height          =   675
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   8115
      Begin VB.Label lbl_procedimento 
         Alignment       =   2  'Center
         Caption         =   "PROCEDIMENTO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   60
         TabIndex        =   1
         Top             =   240
         Width           =   7995
      End
   End
   Begin RichTextLib.RichTextBox rtxt_mensagem 
      Height          =   2655
      Left            =   60
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2940
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   4683
      _Version        =   393217
      ReadOnly        =   -1  'True
      RightMargin     =   9,99999e5
      TextRTF         =   $"analizador_tef.frx":000C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Mensagem para o Operador"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   8055
   End
End
Attribute VB_Name = "analizador_tef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lString As String
Dim lCampo001 As Boolean
Dim lCampo003 As String
Dim lCampo009 As String
Dim lCampo010 As String
Dim lCampo012 As String
Dim lCampo022 As Date
Dim lCampo027 As String
Dim lCampo028 As Long
Dim lProcessa As Boolean



Dim lMensagem As String
Dim lMensagem29(0 To 600) As String
Dim lQtdMensagem29 As Integer
Dim lHoraInicial As Date
Dim BemaRetorno As Integer
Dim lControleSolicitacao As New cControleSolicitacaoTef
Private Sub cmd_ok_Click()
    If Not fValidaValor(txt_valor.Text) > 0 Then
        MsgBox "Informe o valor a ser cancelado.", vbInformation, "Dados Incorreto!"
        txt_valor.SetFocus
    ElseIf txt_nsu.Text = "" Then
        MsgBox "Informe o número da transação do TEF a ser cancelado.", vbInformation, "Dados Incorreto!"
        txt_nsu.SetFocus
    Else
        frm_cnc.Visible = False
        ChamaSolicitacaoCNC
        Unload Me
    End If
End Sub
Private Sub Form_Activate()
    If lProcessa Then
        lProcessa = False
        Call CriaLogTEF(Date & " " & Time & " analizador_tef.Activate: " & gTipoDocumentoFiscal & " N." & gNumeroCupom & " - Valor=" & gValorRecebido & " - gTefResposta=" & gTefResposta & " - " & rtxt_mensagem.Text)
        ChamaRotinas
    Else
        Call CriaLogTEF(Date & " " & Time & " analizador_tef.Activate: SE EU EXECUTASSE SERIA O ERRO DE REDUNDANCIA.")
    End If
End Sub
Private Sub Form_Load()
    Dim xDados As String
    
    CentraForm Me
    gVersao = "v8.03.23a"
    Me.Caption = Me.Caption & " " & gVersao
    gTefResposta = False
    rtxt_mensagem.Visible = False
    lProcessa = True
    ChamaDrive
    
    gImpBematech = False
    gImpSchalter = False
    gImpMecaf = False
    gImpQuick = False
    gImpElgin = False
    gImpDaruma = False
    If gTipoDocumentoFiscal <> "NFCe" Then
        xDados = ReadINI("CUPOM FISCAL", "Impressora Fiscal", ArqSgpIni)
        If xDados = "BEMATECH" Then
            gImpBematech = True
        ElseIf xDados = "SCHALTER" Then
            gImpSchalter = True
        ElseIf xDados = "MECAF" Then
            gImpMecaf = True
        ElseIf xDados = "QUICK" Then
            gImpQuick = True
        ElseIf xDados = "ELGIN" Then
            gImpElgin = True
        ElseIf xDados = "DARUMA" Then
            gImpDaruma = True
        Else
            gImpBematech = True
        End If
    End If
    gLinhasEmBloco = Val(ReadINI("CUPOM FISCAL", "Quantidade de Linhas em Bloco", ArqSgpIni))
    gNomeEmpresa = Val(ReadINI("CUPOM FISCAL", "Nome da Empresa", ArqSgpIni))
    If gLinhasEmBloco = 0 Then
        gLinhasEmBloco = 3
    ElseIf gLinhasEmBloco > 12 Then
        gLinhasEmBloco = 12
    End If
    If gTipoDocumentoFiscal = "NFCe" Then
        gLinhasEmBloco = 1
    End If
End Sub
Function ImprimeTEF(ByVal xTipoImpressao As String) As Boolean
    Dim i As Integer
    Dim i2 As Integer
    Dim xQtdLinha As Integer
    Dim xVias As Integer
    Dim xValorRecebido As String
    Dim xMensagemOld As String
    Dim xString As String
    Dim xACK As Integer
    Dim xST1 As Integer
    Dim xST2 As Integer
    Dim xMensagem29(0 To 200) As String
    
    ImprimeTEF = False
    
    
    xValorRecebido = Mid(Format(gValorRecebido, "000000000000.00"), 1, 12) & Mid(Format(gValorRecebido, "000000000000.00"), 14, 2)
    
    i2 = 0
    xQtdLinha = 0
    If gImpQuick Or gImpDaruma Or gImpBematech Or gTipoDocumentoFiscal = "NFCe" Then
        If xTipoImpressao = "LeituraX" Then
            'Carrega Mensagem
            If Not CarregaMensagemTEF(False, True, False) Then
                'Exit Function
            End If
        End If
        For i = 0 To lQtdMensagem29
            xQtdLinha = xQtdLinha + 1
            xString = Space(48)
            Mid(xString, 1, Len(lMensagem29(i))) = lMensagem29(i)
            xMensagem29(i2) = xMensagem29(i2) & xString
            If xQtdLinha = gLinhasEmBloco Then
                xQtdLinha = 0
                i2 = i2 + 1
            End If
            lMensagem29(i) = ""
        Next
        lQtdMensagem29 = i2
        For i = 0 To lQtdMensagem29
            lMensagem29(i) = xMensagem29(i)
        Next
    End If
    
    
    'Bloqueia Teclado
    'Nao teste qual impressora fiscal
    'pois em ambos os casos uso dll da bematech para travar o teclado
    If gTipoDocumentoFiscal <> "NFCe" Then
        DoEvents
        Call Bematech_FI_IniciaModoTEF
        DoEvents
    End If

    
    'Abre Relatorio Gerencial
    If xTipoImpressao = "LeituraX" Then
        If gImpBematech Then
            BemaRetorno = Bematech_FI_FechaRelatorioGerencial
            BemaRetorno = Bematech_FI_RelatorioGerencial(" ")
        ElseIf gImpQuick Then
            If EcfQuickLeRegistrador("EstadoFiscal", "Long", 5) = "" Then
                AguardaTempo (5)
            End If
            If EcfQuickLeRegistrador("EstadoFiscal", "Long", 5) = "32" Then
                Call EcfQuickEncerraDocumento(0, "Gerencial")
                'Call EcfQuickEncerraDocumento(0, "Gerencial")
            End If
            If EcfQuickLeRegistrador("EstadoFiscal", "Long", 5) = "64" Then
                Call EcfQuickEncerraDocumento("", "Cerrado Informatica")
                'Call EcfQuickEncerraDocumento(0, "Gerencial")
            End If
            If EcfQuickLeRegistrador("EstadoFiscal", "Long", 5) = "" Then
                AguardaTempo (5)
            End If
            If EcfQuickLeRegistrador("EstadoFiscal", "Long", 5) = "32" Then
                Call EcfQuickEncerraDocumento(0, "Gerencial")
                'Call EcfQuickEncerraDocumento(0, "Gerencial")
            End If
            If EcfQuickLeRegistrador("EstadoFiscal", "Long", 5) = "64" Then
                Call EcfQuickEncerraDocumento("", "Cerrado Informatica")
                'Call EcfQuickEncerraDocumento(0, "Gerencial")
            End If
            If EcfQuickDefineGerencial(0, "Gerencial") Then
                BemaRetorno = 1
                If EcfQuickAbreGerencial(0, "Gerencial") Then
                    BemaRetorno = 1
                Else
                    BemaRetorno = 0
                End If
            Else
                BemaRetorno = 0
            End If
        ElseIf gImpElgin Then
            BemaRetorno = Elgin_FechaRelatorioGerencial
            BemaRetorno = Elgin_AbreRelatorioGerencialMFD("01")
        ElseIf gImpDaruma Then
            BemaRetorno = Daruma_TEF_FechaRelatorio()
            BemaRetorno = Daruma_FI_AbreRelatorioGerencial()
        End If
    Else
        If gImpBematech Then
            If UCase(gBandeira) = "TECBAN" Then
                BemaRetorno = Bematech_FI_AbreComprovanteNaoFiscalVinculado("Cartao TecBan   ", xValorRecebido, CStr(Format(gNumeroCupom, "000000")))
            ElseIf UCase(gBandeira) = "TCSMART" Then
                BemaRetorno = Bematech_FI_AbreComprovanteNaoFiscalVinculado("Ticket Car Smart", xValorRecebido, CStr(Format(gNumeroCupom, "000000")))
            Else
                BemaRetorno = Bematech_FI_AbreComprovanteNaoFiscalVinculado("Cartao          ", xValorRecebido, CStr(Format(gNumeroCupom, "000000")))
                Call CriaLogTEF(Date & " " & Time & " Bematech_FI_AbreComprovanteNaoFiscalVinculado: xValorRecebido=" & xValorRecebido & " - gNumeroCupom=" & gNumeroCupom & " - BemaRetorno=" & BemaRetorno)
            End If
        ElseIf gImpQuick Then
            If gConsultaCheque Then
                If EcfQuickAbreCreditoDebito("CONSULTA CHEQUE", gValorRecebido) Then
                    BemaRetorno = 1
                Else
                    BemaRetorno = -1
                End If
            Else
                If EcfQuickAbreCreditoDebito("TEF", gValorRecebido) Then
                    BemaRetorno = 1
                Else
                    BemaRetorno = -1
                End If
            End If
        ElseIf gImpElgin Then
            If gConsultaCheque Then
                BemaRetorno = Elgin_AbreComprovanteNaoFiscalVinculadoMFD("CONSULTA CHEQUE", xValorRecebido, CStr(Format(gNumeroCupom, "000000")), "", "", "")
            Else
                BemaRetorno = Elgin_AbreComprovanteNaoFiscalVinculadoMFD("TEF", xValorRecebido, CStr(Format(gNumeroCupom, "000000")), "", "", "")
            End If
        ElseIf gImpDaruma Then
            xValorRecebido = Format(gValorRecebido, "000000000000.00")
            If gConsultaCheque Then
                BemaRetorno = Daruma_FI_AbreComprovanteNaoFiscalVinculado("CONSULTA CHEQUE", xValorRecebido, CStr(Format(gNumeroCupom, "000000")))
            Else
                'BemaRetorno = Daruma_FI_AbreComprovanteNaoFiscalVinculado("TEF", xValorRecebido, CStr(Format(gNumeroCupom, "000000")))
                BemaRetorno = Daruma_FI_AbreComprovanteNaoFiscalVinculado("Cartao Credito", xValorRecebido, CStr(Format(gNumeroCupom, "000000")))
            End If
        ElseIf gTipoDocumentoFiscal = "NFCe" Then
            ImpTermicaAbreRelatorio
            'Call ImpTermicaImprimeDados("------------------------------------------------", True)
            BemaRetorno = 1
        End If
    End If
    'Verifica se Imprimiu
    Call CriaLogTEF(Date & " " & Time & " ImprimeTEF: " & rtxt_mensagem.Text & " - TipoImpressao=" & xTipoImpressao & " - Quantidade de Vias=" & gQtdViasTEF & " - Quantidade de Linhas de Mensagem=" & lQtdMensagem29)
    If BemaRetorno = 1 Then
        For xVias = 1 To gQtdViasTEF
            For i = 0 To lQtdMensagem29
                'Imprime Texto do TEF
                If xTipoImpressao = "LeituraX" Then
                    If gImpBematech Then
                        'Posto Esmeralda, imprimir o nome do funcionario
                        If gNomeEmpresa Like "*MARQUES DE CASTRO*" And i = 0 Then
                            BemaRetorno = Bematech_FI_RelatorioGerencial(gObservacao1)
                            MsgBox "gObservacao1= ->" & gObservacao1 & "<-"
                        End If
                        BemaRetorno = Bematech_FI_RelatorioGerencial(lMensagem29(i))
                    ElseIf gImpQuick Then
                        If EcfQuickImprimeTexto(lMensagem29(i)) Then
                            BemaRetorno = 1
                        Else
                            BemaRetorno = -1
                        End If
                    ElseIf gImpElgin Then
                        BemaRetorno = Elgin_UsaRelatorioGerencialMFD(lMensagem29(i))
                    ElseIf gImpDaruma Then
                        BemaRetorno = Daruma_FI_RelatorioGerencial(lMensagem29(i))
                    End If
                Else
                    'teste para arrumar linhas
                    If Len(lMensagem29(i)) < 48 Then
                        xString = lMensagem29(i)
                        lMensagem29(i) = Space(48)
                        Mid(lMensagem29(i), 1, Len(xString)) = xString
                    End If
                    If gImpBematech Then
                        'Posto Esmeralda, imprimir o nome do funcionario
                        If gNomeEmpresa Like "*MARQUES DE CASTRO*" And i = 0 Then
                            BemaRetorno = Bematech_FI_RelatorioGerencial(gObservacao1)
                            MsgBox "gObservacao1= ->" & gObservacao1 & "<-"
                        End If
                        BemaRetorno = Bematech_FI_UsaComprovanteNaoFiscalVinculado(lMensagem29(i))
                    ElseIf gImpQuick Then
                        If EcfQuickImprimeTexto(lMensagem29(i)) Then
                            BemaRetorno = 1
                        Else
                            BemaRetorno = -1
                        End If
                    ElseIf gImpElgin Then
                        BemaRetorno = Elgin_UsaComprovanteNaoFiscalVinculado(lMensagem29(i))
                    ElseIf gImpDaruma Then
                        BemaRetorno = Daruma_FI_UsaComprovanteNaoFiscalVinculado(lMensagem29(i))
                    ElseIf gTipoDocumentoFiscal = "NFCe" Then
                        'MsgBox "i: " & i & " - xVias: " & xVias & " - Texto Mid: " & Mid(lMensagem29(i), 1, 11) & " Mensagem - " & lMensagem29(i)
                        If i = 1 Then
                            If Mid(lMensagem29(i), 1, 11) = "Funcionario" Or Mid(lMensagem29(i), 1, 11) = "Funcionário" Then
                                If xVias = 1 Then
                                    Call ImpTermicaImprimeDados("     *****    VIA DA EMPRESA    *****", True)
                                Else
                                    Call ImpTermicaImprimeDados("     *****    VIA DO CLIENTE    *****", True)
                                End If
                            End If
                        End If
                        Call ImpTermicaImprimeDados(lMensagem29(i), True)
                    End If
                End If
                'Verifica se Não Imprimiu
                If gImpBematech Then
                    If BemaRetorno <> 1 Then
                        Exit Function
                    End If
                    BemaRetorno = Bematech_FI_RetornoImpressora(xACK, xST1, xST2)
                    If BemaRetorno <> 1 Then
                        Exit Function
                    End If
                    If xST2 = 1 Then
                        Call CriaLogTEF(Date & " " & Time & " Bematech_FI_RetornoImpressora: Sensor de Papel 2(Guilhotina) - xACK=" & xACK & " - xST1=" & xST1 & " - xST2=" & xST1)
                        Exit Function
                    End If
                    'If xST1 = 64 Then
                    '    Call CriaLogTEF(Date & " " & Time & " Bematech_FI_RetornoImpressora: Sensor de Papel 1(Bobina) - xACK=" & xACK & " - xST1=" & xST1 & " - xST2=" & xST1)
                    '    Exit Function
                    'End If
                ElseIf gImpQuick Then
                    If BemaRetorno <> 1 Then
                        Exit Function
                    End If
                ElseIf gImpElgin Then
                    If BemaRetorno <> 1 Then
                        Exit Function
                    End If
                ElseIf gImpDaruma Then
                    If BemaRetorno <> 1 Then
                        Exit Function
                    End If
                End If
                    
                If (gImpBematech And xVias = 2 And i = 6) Or (gImpQuick And xVias = 2 And i = 0) Or (gImpElgin And xVias = 2 And i = 0) Or (gImpDaruma And xVias = 2 And i = 0) Then
                    'Pausa de 2 Segundos Entre a Primeira e Segunda Via
                    'para Cortar o Papel
                    xMensagemOld = rtxt_mensagem.Text
                    rtxt_mensagem.Text = "Primeira Via Já Está Impressa." & Chr(10) & Chr(10) & "Recorte Agora!"
                    rtxt_mensagem.Visible = True
                    DoEvents
                    lHoraInicial = Time
                    Do Until DateDiff("s", lHoraInicial, Time) >= 2
                        DoEvents
                    Loop
                    'rtxt_mensagem.Text = "Aguarde!" & Chr(10) & Chr(10) & "Imprimindo TEF"
                    rtxt_mensagem.Text = xMensagemOld
                    'Call CriaLogTEF("ImprimeTEF: Mensagem 15: " & rtxt_mensagem.Text & " às: " & Time)
                    rtxt_mensagem.Visible = True
                    DoEvents
                End If
                If xVias = 1 And i = lQtdMensagem29 Then
                    If gImpBematech Then
                        BemaRetorno = Bematech_FI_RelatorioGerencial(Space(144))
                        'Pular mais 2 Linhas antes de cortar a segunda via
                        BemaRetorno = Bematech_FI_RelatorioGerencial(Space(gLinhasEntreCV * 48))
                        BemaRetorno = Bematech_FI_AcionaGuilhotinaMFD(1)
                    ElseIf gImpQuick Then
                        If EcfQuickImprimeTexto(Space(144)) Then
                            BemaRetorno = 1
                        Else
                            BemaRetorno = -1
                        End If
                    ElseIf gImpElgin Then
                        If xTipoImpressao = "LeituraX" Then
                            BemaRetorno = Elgin_UsaRelatorioGerencialMFD(Space(144))
                        Else
                            BemaRetorno = Elgin_UsaComprovanteNaoFiscalVinculado(Space(144))
                        End If
                    ElseIf gImpDaruma Then
                        If xTipoImpressao = "LeituraX" Then
                            BemaRetorno = Daruma_FI_RelatorioGerencial(Space(144))
                        Else
                            BemaRetorno = Daruma_FI_UsaComprovanteNaoFiscalVinculado(Space(144))
                        End If
                    ElseIf gTipoDocumentoFiscal = "NFCe" Then
                        Call ImpTermicaImprimeDados("-------------- recorte aqui --------------------", True)
                        Call ImpTermicaImprimeDados(" ", True)
                        Call ImpTermicaImprimeDados(" ", True)
                        Call ImpTermicaImprimeDados(" ", True)
                        ImpTermicaFechaRelatorio ("Relatório de Notas TEF (Impressora Térica)")
                        If gQtdViasTEF > 1 Then
                            ImpTermicaAbreRelatorio
                        End If
                        ''Call ImpTermicaImprimeDados("------------------------------------------------", True)
                    End If
                End If
            Next
            If gQtdViasTEF = 2 Then
                'Imprime 3 linhas em branco
                If xTipoImpressao = "LeituraX" Then
                    If gImpBematech Then
                        BemaRetorno = Bematech_FI_RelatorioGerencial(Space(96))
                    ElseIf gImpQuick Then
                        If EcfQuickImprimeTexto(Space(96)) Then
                            BemaRetorno = 1
                        Else
                            BemaRetorno = -1
                        End If
                    ElseIf gImpElgin Then
                        BemaRetorno = Elgin_UsaRelatorioGerencialMFD(Space(96))
                    ElseIf gImpDaruma Then
                        BemaRetorno = Daruma_FI_RelatorioGerencial(Space(96))
                    End If
                Else
                    If gImpBematech Then
                        BemaRetorno = Bematech_FI_UsaComprovanteNaoFiscalVinculado(Space(96))
                    ElseIf gImpQuick Then
                        If EcfQuickImprimeTexto(Space(96)) Then
                            BemaRetorno = 1
                        Else
                            BemaRetorno = -1
                        End If
                    ElseIf gImpElgin Then
                        BemaRetorno = Elgin_UsaComprovanteNaoFiscalVinculado(Space(96))
                    ElseIf gImpDaruma Then
                        BemaRetorno = Daruma_FI_UsaComprovanteNaoFiscalVinculado(Space(96))
                    End If
                End If
                'Verifica se Não Imprimiu
                If BemaRetorno <> 1 Then
                    Exit Function
                End If
            End If
        Next
    Else
        Exit Function
    End If
    'Fecha Relatorio Gerencial
    If xTipoImpressao = "LeituraX" Then
        If gImpBematech Then
            BemaRetorno = Bematech_FI_FechaRelatorioGerencial
        ElseIf gImpQuick Then
            If EcfQuickEncerraDocumento("", "Cerrado Informatica") Then
                BemaRetorno = 1
            Else
                BemaRetorno = -1
            End If
        ElseIf gImpElgin Then
            BemaRetorno = Elgin_FechaRelatorioGerencial()
        ElseIf gImpDaruma Then
            BemaRetorno = Daruma_TEF_FechaRelatorio()
        End If
    Else
        If gImpBematech Then
            BemaRetorno = Bematech_FI_FechaComprovanteNaoFiscalVinculado
        ElseIf gImpQuick Then
            If EcfQuickEncerraDocumento("", "Cerrado Informatica") Then
                BemaRetorno = 1
            Else
                BemaRetorno = -1
            End If
        ElseIf gImpElgin Then
            BemaRetorno = Elgin_FechaComprovanteNaoFiscalVinculado()
        ElseIf gImpDaruma Then
            BemaRetorno = Daruma_TEF_FechaRelatorio()
        ElseIf gTipoDocumentoFiscal = "NFCe" Then
            If gQtdViasTEF > 1 Then
                'Call ImpTermicaImprimeDados("------------------------------------------------", True)
                Call ImpTermicaImprimeDados(" ", True)
                Call ImpTermicaImprimeDados(" ", True)
                Call ImpTermicaImprimeDados(" ", True)
                ImpTermicaFechaRelatorio ("Relatório de Notas TEF (Impressora Térica)")
            End If
        End If
    End If
    'Verifica se Imprimiu
    If BemaRetorno = 1 Then
        ImprimeTEF = True
    Else
        Exit Function
    End If
End Function
Function ImprimeTefADM(ByVal xTipoImpressao As String) As Boolean
    Dim i As Integer
    Dim i2 As Integer
    Dim xQtdLinha As Integer
    Dim xVias As Integer
    Dim xMensagemOld As String
    Dim xString As String
    Dim xMensagem29(0 To 200) As String
    
    ImprimeTefADM = False
    
    
    'Abre Relatorio Gerencial
    If gImpBematech Then
        BemaRetorno = Bematech_FI_FechaRelatorioGerencial
        BemaRetorno = Bematech_FI_RelatorioGerencial(" ")
    ElseIf gImpQuick Then
        If EcfQuickLeRegistrador("EstadoFiscal", "Long", 5) = "" Then
            AguardaTempo (5)
        End If
        If EcfQuickLeRegistrador("EstadoFiscal", "Long", 5) = "32" Then
            Call EcfQuickEncerraDocumento(0, "Gerencial")
            'Call EcfQuickEncerraDocumento(0, "Gerencial")
        End If
        If EcfQuickLeRegistrador("EstadoFiscal", "Long", 5) = "64" Then
            Call EcfQuickEncerraDocumento("", "Cerrado Informatica")
            'Call EcfQuickEncerraDocumento(0, "Gerencial")
        End If
        If EcfQuickLeRegistrador("EstadoFiscal", "Long", 5) = "" Then
            AguardaTempo (5)
        End If
        If EcfQuickLeRegistrador("EstadoFiscal", "Long", 5) = "32" Then
            Call EcfQuickEncerraDocumento(0, "Gerencial")
            'Call EcfQuickEncerraDocumento(0, "Gerencial")
        End If
        If EcfQuickLeRegistrador("EstadoFiscal", "Long", 5) = "64" Then
            Call EcfQuickEncerraDocumento("", "Cerrado Informatica")
            'Call EcfQuickEncerraDocumento(0, "Gerencial")
        End If
        If EcfQuickDefineGerencial(0, "Gerencial") Then
            BemaRetorno = 1
            If EcfQuickAbreGerencial(0, "Gerencial") Then
                BemaRetorno = 1
            Else
                BemaRetorno = 0
            End If
        Else
            BemaRetorno = 0
        End If
    ElseIf gImpElgin Then
        BemaRetorno = Elgin_FechaRelatorioGerencial
        BemaRetorno = Elgin_AbreRelatorioGerencialMFD("01")
    ElseIf gImpDaruma Then
        BemaRetorno = Daruma_TEF_FechaRelatorio
        BemaRetorno = Daruma_FI_AbreRelatorioGerencial()
    ElseIf gTipoDocumentoFiscal = "NFCe" Then
        ImpTermicaAbreRelatorio
        'Call ImpTermicaImprimeDados("------------------------------------------------", True)
        BemaRetorno = 1
    End If
    
    'Bloqueia Teclado
    'Nao teste qual impressora fiscal
    'pois em ambos os casos uso dll da bematech para travar o teclado
    If gTipoDocumentoFiscal <> "NFCe" Then
        DoEvents
        Call Bematech_FI_IniciaModoTEF
        DoEvents
    End If
    
    'Verifica se Imprimiu
    Call CriaLogTEF(Date & " " & Time & " ImprimeTefADM: " & rtxt_mensagem.Text & " - TipoImpressao=" & xTipoImpressao & " - Quantidade de Vias=" & gQtdViasTEF & " - Quantidade de Linhas de Mensagem=" & lQtdMensagem29)
    If BemaRetorno = 1 Then
    
        'ECF Quick
        i2 = 0
        xQtdLinha = 0
        If gImpQuick Or gImpElgin Or gImpDaruma Or gImpBematech Or gTipoDocumentoFiscal = "NFCe" Then
            If xTipoImpressao = "LeituraX" Then
                'Carrega Mensagem
                If Not CarregaMensagemTEF(False, True, False) Then
                    'Exit Function
                End If
            End If
            For i = 0 To lQtdMensagem29
                xQtdLinha = xQtdLinha + 1
                xString = Space(48)
                Mid(xString, 1, Len(lMensagem29(i))) = lMensagem29(i)
                xMensagem29(i2) = xMensagem29(i2) & xString
                If xQtdLinha = gLinhasEmBloco Then
                    xQtdLinha = 0
                    i2 = i2 + 1
                End If
                lMensagem29(i) = ""
            Next
            lQtdMensagem29 = i2
            For i = 0 To lQtdMensagem29
                lMensagem29(i) = xMensagem29(i)
            Next
        End If
    
    
        For xVias = 1 To gQtdViasTEF
            For i = 0 To lQtdMensagem29
                'Imprime Texto do TEF
                If gImpBematech Then
                    'xString = Space(48)
                    'Mid(xString, 1, Len(lMensagem29(i))) = lMensagem29(i)
                    'BemaRetorno = Bematech_FI_RelatorioGerencial(xString)
                    BemaRetorno = Bematech_FI_RelatorioGerencial(lMensagem29(i))
                ElseIf gImpQuick Then
                    If EcfQuickImprimeTexto(lMensagem29(i)) Then
                        BemaRetorno = 1
                    Else
                        BemaRetorno = -1
                    End If
                ElseIf gImpElgin Then
                    BemaRetorno = Elgin_UsaRelatorioGerencialMFD(lMensagem29(i))
                ElseIf gImpDaruma Then
                    BemaRetorno = Daruma_FI_RelatorioGerencial(lMensagem29(i))
                ElseIf gTipoDocumentoFiscal = "NFCe" Then
                    Call ImpTermicaImprimeDados(lMensagem29(i), True)
                End If
                'Verifica se Não Imprimiu
                If BemaRetorno <> 1 Then
                    Exit Function
                End If
                If xVias = 2 And i = 6 Then
                    'Pausa de 2 Segundos Entre a Primeira e Segunda Via
                    'para Cortar o Papel
                    xMensagemOld = rtxt_mensagem.Text
                    rtxt_mensagem.Text = "Primeira Via Já Está Impressa." & Chr(10) & Chr(10) & "Recorte Agora!"
                    'Call CriaLogTEF("ImprimeTefADM: Mensagem 16: " & rtxt_mensagem.Text & " às: " & Time)
                    rtxt_mensagem.Visible = True
                    DoEvents
                    lHoraInicial = Time
                    Do Until DateDiff("s", lHoraInicial, Time) >= 2
                        DoEvents
                    Loop
                    'rtxt_mensagem.Text = "Aguarde!" & Chr(10) & Chr(10) & "Imprimindo TEF" & Chr(10) & "Administrativo."
                    rtxt_mensagem.Text = xMensagemOld
                    'Call CriaLogTEF("ImprimeTefADM: Mensagem 17: " & rtxt_mensagem.Text & " às: " & Time)
                    rtxt_mensagem.Visible = True
                    DoEvents
                End If
            Next
            If xVias = 1 Then
                'Imprime 2 linhas em branco
                If gImpBematech Then
                    BemaRetorno = Bematech_FI_RelatorioGerencial(Space(144))
                ElseIf gImpQuick Then
                    If EcfQuickImprimeTexto(Space(144)) Then
                        BemaRetorno = 1
                    Else
                        BemaRetorno = -1
                    End If
                ElseIf gImpElgin Then
                    BemaRetorno = Elgin_UsaRelatorioGerencialMFD(Space(144))
                ElseIf gImpDaruma Then
                    BemaRetorno = Daruma_FI_RelatorioGerencial(Space(144))
                    ElseIf gTipoDocumentoFiscal = "NFCe" Then
                        Call ImpTermicaImprimeDados("-------------- recorte aqui --------------------", True)
                        Call ImpTermicaImprimeDados(" ", True)
                        Call ImpTermicaImprimeDados(" ", True)
                        Call ImpTermicaImprimeDados(" ", True)
                        ImpTermicaFechaRelatorio ("Relatório de Notas TEF (Impressora Térica)")
                        If gQtdViasTEF > 1 Then
                            ImpTermicaAbreRelatorio
                        End If
                        ''Call ImpTermicaImprimeDados("------------------------------------------------", True)
                End If
                'Verifica se Não Imprimiu
                If BemaRetorno <> 1 Then
                    Exit Function
                End If
            End If
        Next
    Else
        Exit Function
    End If
    'Fecha Relatorio Gerencial
    If gImpBematech Then
        BemaRetorno = Bematech_FI_FechaRelatorioGerencial
    ElseIf gImpQuick Then
        If EcfQuickEncerraDocumento("", "Cerrado Informatica") Then
            BemaRetorno = 1
        Else
            BemaRetorno = -1
        End If
    ElseIf gImpElgin Then
        BemaRetorno = Elgin_FechaRelatorioGerencial()
    ElseIf gImpDaruma Then
        BemaRetorno = Daruma_TEF_FechaRelatorio()
    ElseIf gTipoDocumentoFiscal = "NFCe" Then
        If gQtdViasTEF > 1 Then
            'Call ImpTermicaImprimeDados("------------------------------------------------", True)
            Call ImpTermicaImprimeDados(" ", True)
            Call ImpTermicaImprimeDados(" ", True)
            Call ImpTermicaImprimeDados(" ", True)
            ImpTermicaFechaRelatorio ("Relatório de Notas TEF (Impressora Térica)")
        End If
    End If
    'Verifica se Imprimiu
    If BemaRetorno = 1 Then
        ImprimeTefADM = True
    Else
        Exit Function
    End If
End Function
Function AtivaGerenciadorPadrao() As Boolean
    Dim retval As Long
    AtivaGerenciadorPadrao = False
    If gNomeGerenciadorPadrao <> "" Then
        retval = Shell(gNomeGerenciadorPadrao, vbMinimizedNoFocus)
    End If
'    If UCase(gBandeira) = "TECBAN" Then
'        RetVal = Shell("C:\tef_disc\tef_disc.exe", vbMinimizedNoFocus)
'    ElseIf UCase(gBandeira) = "PAGCARD" Then
'        RetVal = Shell("C:\CardTech_NEUS\SAC.exe", vbMinimizedNoFocus)
'    Else
'        RetVal = Shell("C:\tef_dial\tef_dial.exe", vbMinimizedNoFocus)
'    End If
End Function
Function CarregaMensagemTEF(ByVal pMostraMensagemNova As Boolean, ByVal pMostraMensagemAntiga As Boolean, ByVal pTemporizarMensagem As Boolean) As Boolean
    Dim i As Integer
    Dim i2 As Integer
    Dim xMensagemAntiga
    Dim xNomeArquivoResp As String 'Arquivo de Resposta
    Dim xUsarCampo715 As Boolean
    
    On Error GoTo FileError
    
    CarregaMensagemTEF = False
'    If UCase(gBandeira) = "TECBAN" Then
'        xNomeArquivoResp = "C:\TEF_DISC\RESP\IntPos.001"
'    ElseIf UCase(gBandeira) = "TCSMART" Then
'        xNomeArquivoResp = "C:\TCS\RX\IntTCS.001"
'    ElseIf UCase(gBandeira) = "SMARTEF" Then
'        xNomeArquivoResp = "C:\SMARTEF\RESP\IntPos.001"
'    ElseIf UCase(gBandeira) = "SUPERTEF" Then
'        xNomeArquivoResp = "C:\SUPERTEF\RESP\IntPos.001"
'    ElseIf UCase(gBandeira) = "HIPERTEF" Or UCase(gBandeira) = "HIPERCARD" Then
'        xNomeArquivoResp = "C:\HiperTEF\RESP\IntPos.001"
'    ElseIf UCase(gBandeira) = "PAGCARD" Then
'        xNomeArquivoResp = "C:\CardTech_NEUS\RESP\IntPos.001"
'    Else
'        xNomeArquivoResp = "C:\TEF_DIAL\RESP\IntPos.001"
'    End If
    If UCase(gBandeira) = "TCSMART" Then
        xNomeArquivoResp = gDiretorioResp & "IntTCS.001"
    Else
        xNomeArquivoResp = gDiretorioResp & "IntPos.001"
    End If
    lCampo001 = False
    lCampo003 = ""
    lCampo009 = ""
    lCampo010 = ""
    lCampo012 = ""
    lCampo022 = "00:00:00"
    lCampo027 = ""
    lCampo028 = 0
    xUsarCampo715 = False
   
    
    If gArqTxt.FileExists(xNomeArquivoResp) Then
        'BUG:XP DÁ UM TEMPO PARA LIBERAR ARQUIVO PARA LEITURA
        lHoraInicial = Time
        Do Until DateDiff("s", lHoraInicial, Time) >= 2
        Loop
        Call CopiarArquivo(xNomeArquivoResp, "c:\vb5\sgp\data\teste.txt")
    
        'Teste para descobrir Bug de voltar a pedir
        'cartao após imprimir.
        Dim xNomeArquivoCopia As String
        ' 123456789012345678901234567890
        '"TTF_bug_dd_MM_yyyy__HH:mm:ss.LOG"
        xNomeArquivoCopia = "TTF_bug_" & Format(Date, "dd") & "_" & Format(Date, "MM") & "_" & Format(Date, "yyyy") & "__" & Format(Time, "HH:mm:ss") & ".LOG"
        Mid(xNomeArquivoCopia, 23, 1) = "_"
        Mid(xNomeArquivoCopia, 26, 1) = "_"
        xNomeArquivoCopia = "c:\vb5\sgp\data\" & xNomeArquivoCopia
        Call CopiarArquivo(xNomeArquivoResp, xNomeArquivoCopia)
        'fim do teste acima
        
        Set gArquivo = gArqTxt.OpenTextFile(xNomeArquivoResp, ForReading)
        lMensagem = ""
        i = 0
        lQtdMensagem29 = -1
        
        'Pega gTextoAntesCV e colocar como texto a ser impresso
        If Len(gTextoAntesCV) > 0 Then
            lQtdMensagem29 = (Len(gTextoAntesCV) / 48)
            For i = 1 To lQtdMensagem29
                lMensagem29(i - 1) = Mid(gTextoAntesCV, (i * 48 - 47), 48)
            Next
            lQtdMensagem29 = lQtdMensagem29 - 1
        End If
        
        Do Until gArquivo.AtEndOfStream
            lString = gArquivo.ReadLine
            If UCase(gBandeira) = "TCSMART" Then
                If Mid(lString, 1, 7) = "501-000" Then
                    If Mid(lString, 11, 10) = gNumeroControleSolicitacao Then
                        lCampo001 = True
                        lCampo009 = "0"
                    End If
                End If
                If Mid(lString, 1, 3) = "515" Then
                    i2 = Len(lString)
                    lQtdMensagem29 = lQtdMensagem29 + 1
                    lMensagem29(lQtdMensagem29) = Mid(lString, 12, i2 - 12)
                    If lMensagem29(lQtdMensagem29) = "" Then
                        lMensagem29(lQtdMensagem29) = " "
                    End If
                End If
                If Mid(lString, 1, 7) = "514-000" Then
                    i = i + 1
                    If i > 1 Then
                        lMensagem = lMensagem & Chr(10)
                    End If
                    lCampo028 = Val(Mid(lString, 11, Len(lString) - 10))
                    lMensagem = lMensagem & Mid(lString, 11, Len(lString) - 10) & " Linhas à Imprimir"
                End If
            Else
                If Mid(lString, 1, 7) = "001-000" Then
                    If Mid(lString, 11, 10) = gNumeroControleSolicitacao Then
                        lCampo001 = True
                    End If
                End If
                If Mid(lString, 1, 7) = "003-000" Then
                    If Len(lString) >= 11 Then
                        'If gBandeira = "VISANET" Then
                            lCampo003 = Format(CCur(Mid(lString, 11, Len(lString) - 10)) / 100, "###,###,##0.00")
                        'Else
                        '    lCampo003 = Mid(lString, 11, Len(lString) - 10)
                        'End If
                    End If
                End If
                If Mid(lString, 1, 7) = "008-000" Then
                    Call CriaLogTEF(Date & " " & Time & " CarregaMensagemTEF: Valor de Desconto Concedido = " & CCur(Mid(lString, 11, Len(lString) - 10)) / 100)
                    gValorDescontoConcedido = Mid(lString, 11, Len(lString) - 10)
                    Call CriaLogTEF(Date & " " & Time & " CarregaMensagemTEF: gValorDescontoConcedido = " & gValorDescontoConcedido / 100)
                End If
                If Mid(lString, 1, 7) = "009-000" Then
                    lCampo009 = Mid(lString, 11, Len(lString) - 10)
                    If lCampo009 = "00" And UCase(gBandeira) = "SMARTEF" Then
                        lCampo009 = "0"
                    End If
                End If
                If Mid(lString, 1, 7) = "010-000" Then
                    If Len(lString) >= 11 Then
                        lCampo010 = Mid(lString, 11, Len(lString) - 10)
                    End If
                End If
                If Mid(lString, 1, 7) = "012-000" Then
                    lCampo012 = Mid(lString, 11, Len(lString) - 10)
                End If
                If Mid(lString, 1, 7) = "016-000" Then
                    If Len(lString) >= 11 Then
                        Call CriaLogTEF(Date & " " & Time & " CarregaMensagemTEF: " & rtxt_mensagem.Text & " - DataCV 016-000 Inicial=" & lCampo022)
                        lCampo022 = CDate(Mid(lString, 11, 2) & "/" & Mid(lString, 13, 2) & "/" & Format(Date, "yyyy"))
                        Call CriaLogTEF(Date & " " & Time & " CarregaMensagemTEF: " & rtxt_mensagem.Text & " - DataCV 016-000 Final=" & lCampo022)
                    End If
                End If
                If Mid(lString, 1, 7) = "022-000" Then
                    If Len(lString) >= 11 Then
                        Call CriaLogTEF(Date & " " & Time & " CarregaMensagemTEF: " & rtxt_mensagem.Text & " - DataCV 022-000 Inicial=" & lCampo022)
                        If Not IsDate(Mid(lString, 11, 2) & "/" & Mid(lString, 13, 2) & "/" & Mid(lString, 15, 4)) Then
                            lCampo022 = Date
                        Else
                            lCampo022 = CDate(Mid(lString, 11, 2) & "/" & Mid(lString, 13, 2) & "/" & Mid(lString, 15, 4))
                        End If
                        Call CriaLogTEF(Date & " " & Time & " CarregaMensagemTEF: " & rtxt_mensagem.Text & " - DataCV 022-000 Final=" & lCampo022)
                    End If
                End If
                If Mid(lString, 1, 7) = "027-000" Then
                    If Len(lString) >= 11 Then
                        lCampo027 = Mid(lString, 11, Len(lString) - 10)
                    End If
                End If
                If Mid(lString, 1, 7) = "028-000" Then
                    lCampo028 = Mid(lString, 11, Len(lString) - 10)
                End If
                If Mid(lString, 1, 3) = "029" Then
                    i2 = Len(lString)
                    lQtdMensagem29 = lQtdMensagem29 + 1
                    lMensagem29(lQtdMensagem29) = Mid(lString, 12, i2 - 12)
                    If lMensagem29(lQtdMensagem29) = "" Then
                        lMensagem29(lQtdMensagem29) = " "
                    End If
                End If
                If Mid(lString, 1, 3) = "030" Then
                    i = i + 1
                    If i > 1 Then
                        lMensagem = lMensagem & Chr(10)
                    End If
                    lMensagem = lMensagem & Mid(lString, 11, Len(lString) - 10)
                End If
                If Mid(lString, 1, 7) = "714-000" Then
                    Call CriaLogTefEspecial("CarregaMensagemTEF - 101 Testa campo 714-000. E verifica se lCampo028 = 0. lCampo028=" & lCampo028)
                    If lCampo028 = 0 Then
                        Call CriaLogTefEspecial("CarregaMensagemTEF - 102 Redefine lCampo028 e xUsarCampo715 = True.")
                        lCampo028 = Mid(lString, 11, Len(lString) - 10)
                        xUsarCampo715 = True
                    End If
                End If
                If Mid(lString, 1, 3) = "715" Then
                    Call CriaLogTefEspecial("CarregaMensagemTEF - 103 Monta mensagem com campo 715. xUsarCampo715: " & xUsarCampo715)
                    If xUsarCampo715 = True Then
                        i2 = Len(lString)
                        lQtdMensagem29 = lQtdMensagem29 + 1
                        lMensagem29(lQtdMensagem29) = Mid(lString, 12, i2 - 12)
                        If lMensagem29(lQtdMensagem29) = "" Then
                            lMensagem29(lQtdMensagem29) = " "
                        End If
                    End If
                End If
            End If
        Loop
        If pMostraMensagemNova Then
            xMensagemAntiga = rtxt_mensagem.Text
            rtxt_mensagem.Text = lMensagem
            'Call CriaLogTEF("CarregaMensagemTEF: Mensagem 18: " & rtxt_mensagem.Text & " às: " & Time)
            rtxt_mensagem.Visible = True
            If (lCampo009 = "0" Or lCampo009 = "P1") And pTemporizarMensagem = True Then
                lHoraInicial = Time
                Do Until DateDiff("s", lHoraInicial, Time) >= 10
                    DoEvents
                Loop
                If lCampo010 = "REDECARD" Then
                    lHoraInicial = Time
                    Do Until DateDiff("s", lHoraInicial, Time) >= 10
                        DoEvents
                    Loop
                End If
            End If
            'rtxt_mensagem.Visible = False
            'rtxt_mensagem.Text = xMensagemAntiga
            'Call CriaLogTEF("CarregaMensagemTEF: Mensagem 19: " & rtxt_mensagem.Text & " às: " & Time)
            'rtxt_mensagem.Visible = True
            DoEvents
        End If
        gArquivo.Close
    Else
        MsgBox "O arquivo IntPos.001 foi deletado de forma não identificada!", vbInformation, "Mensagem Padrão"
        Exit Function
    End If
    'If lQtdMensagem29 <> -1 Then
        CarregaMensagemTEF = True
    'End If
    Exit Function

FileError:
    Call CriaLogTEF(Date & " " & Time & " CarregaMensagemTEF: " & " - ErroNúmero: " & Err & " - ErroTexto: " & Error)
    MsgBox "ERRO não identificado!", vbInformation, "Rotina: CarregaMensagemTEF"
    Exit Function

End Function
Private Function CopiaRegistrada() As Boolean
    Dim xString As String
    Dim xDataBase As Date
    Dim xNumeroSerieEcf As String
    Dim xDataEcf As Date
    Dim xDataMicro As Date
    Dim xFaseErro As Integer
    
    On Error GoTo FileError
    
    CopiaRegistrada = False
    xFaseErro = 1
    xDataBase = CDate("25/12/2020")
    xFaseErro = 2
    If gArqTxt.FileExists("C:\WINDOWS\SYSTEM\Msretc.dep") Then
        xFaseErro = 3
        Set gArquivo = gArqTxt.OpenTextFile("C:\WINDOWS\SYSTEM\Msretc.dep", ForReading)
        xFaseErro = 4
        xString = gArquivo.ReadLine
        xFaseErro = 5
        xNumeroSerieEcf = xString
        xFaseErro = 6
        xString = gArquivo.ReadLine
        xFaseErro = 7
        xDataEcf = CDate(Mid(xString, 1, 10))
        xFaseErro = 8
        xString = gArquivo.ReadLine
        xFaseErro = 9
        xDataMicro = CDate(Mid(xString, 1, 10))
        xFaseErro = 10
        gArquivo.Close
        xFaseErro = 11
    Else
        xFaseErro = 12
        Call CriaLogTEF(Date & " " & Time & " CopiaRegistrada Arquivo não encontrado: ")
    End If
    
    xFaseErro = 13
    If VerificaNSImpressoraFiscal(xDataBase) Then
        xFaseErro = 14
        Call CriaLogTEF(Date & " " & Time & " CopiaRegistrada OK: " & rtxt_mensagem.Text & " - NumeroSerieEcf=" & xNumeroSerieEcf)
        xFaseErro = 15
        CopiaRegistrada = True
        Exit Function
    End If
    
    xFaseErro = 16
    If xDataEcf < xDataBase Then
        xFaseErro = 17
        Call CriaLogTEF(Date & " " & Time & " CopiaRegistrada OK: " & rtxt_mensagem.Text & " - DataECF=" & xDataEcf & " - DataMicro=" & xDataMicro & " - DataBase=" & xDataBase)
        xFaseErro = 18
        If DateDiff("d", xDataMicro, Date) < 2 Then
            xFaseErro = 19
            CopiaRegistrada = True
            Exit Function
        End If
    End If
    xFaseErro = 20
    Call CriaLogTEF(Date & " " & Time & " CopiaRegistrada Restrição: " & rtxt_mensagem.Text & " - DataECF=" & xDataEcf & " - DataMicro=" & xDataMicro & " - DataBase=" & xDataBase & " - NumeroSerieEcf=" & xNumeroSerieEcf)
    Exit Function

FileError:
    Call CriaLogTEF(Date & " " & Time & " Erro CopiaRegistrada: " & " - ErroNúmero: " & Err & " - FaseErro: " & xFaseErro & " - ErroTexto: " & Error)
    MsgBox "ERRO não identificado!", vbInformation, "Rotina: CopiaRegistrada"
    Exit Function
End Function
Private Sub ChamaImprimeTEF()
    
    
    
    'gTefString = "ImprimeTEF"
    'ElseIf gTefString = "ImprimeTEF" Then
    rtxt_mensagem.Text = "Aguarde!" & Chr(10) & Chr(10) & "Imprimindo TEF."
    'Call CriaLogTEF("ChamaImprimeTEF: Mensagem 20: " & rtxt_mensagem.Text & " às: " & Time)
    rtxt_mensagem.Visible = True
    DoEvents
    gTefResposta = TestaImprimeTEF("Vinculado")
    rtxt_mensagem.Visible = False
    If gTefResposta Then
        '''ImprimeTEF = True
        'MsgBox "Confirma CNF"
        'gTefString = "CNF"
        'ElseIf gTefString = "CNF" Then
        rtxt_mensagem.Text = "Aguarde!" & Chr(10) & Chr(10) & "Confirmando TEF."
        'Call CriaLogTEF("ChamaImprimeTEF: Mensagem 21: " & rtxt_mensagem.Text & " às: " & Time)
        rtxt_mensagem.Visible = True
        DoEvents
        gTefResposta = CNF
        rtxt_mensagem.Visible = False
    Else
        'MsgBox "Cancela NCN"
        'gTefString = "NCN"
        'ElseIf gTefString = "NCN" Then
        rtxt_mensagem.Text = "Aguarde!" & Chr(10) & Chr(10) & "Cancelando TEF."
        'Call CriaLogTEF("ChamaImprimeTEF: Mensagem 22: " & rtxt_mensagem.Text & " às: " & Time)
        rtxt_mensagem.Visible = True
        DoEvents
        gTefResposta = SolicitacaoNCN
        rtxt_mensagem.Visible = False
    End If
    

End Sub
Private Sub ChamaRotinas()
    Call CriaLogTEF(Date & " " & Time & " ChamaRotinas: " & gTefString)
    If gTefString = "SolicitacaoADM" Then
        lbl_procedimento.Caption = "Funções Administrativas"
        ChamaSolicitacaoADM
    ElseIf gTefString = "SolicitacaoAlteraPrecoTCS" Then
        lbl_procedimento.Caption = "Função Altera Preço TCS"
        ChamaSolicitacaoAlteraPrecoTCS
    ElseIf gTefString = "SolicitacaoDSC" Then
        lbl_procedimento.Caption = "Pedido de Autorização de Desconto"
        ChamaSolicitacaoDSC
    ElseIf gTefString = "SolicitacaoTEF" Then
        lbl_procedimento.Caption = "Pedido de Autorização por Cartão"
        ChamaSolicitacaoTEF
    ElseIf gTefString = "SolicitacaoCNC" Then
        frm_cnc.Visible = True
        lbl_procedimento.Caption = "Cancelamento de Venda"
        txt_valor.SetFocus
        Exit Sub
    ElseIf gTefString = "SolicitacaoConsultaCH" Then
        lbl_procedimento.Caption = "Consulta Cheque Serasa"
        ChamaSolicitacaoConsultaCH
    ElseIf gTefString = "SolicitacaoNCN" Then
        lbl_procedimento.Caption = "Não Confirmação de Venda/Impressão"
        ChamaSolicitacaoNCN
    ElseIf gTefString = "ImprimeTEF" Then
        lbl_procedimento.Caption = "Impressão do TEF"
        ChamaImprimeTEF
    End If
    Unload Me
End Sub
Private Sub ChamaSolicitacaoADM()
    Dim xHoraInicial As Date
    
    
    'gTefString = "GerenciadorPadraoAtivo"
    'If gTefString = "GerenciadorPadraoAtivo" Then
    rtxt_mensagem.Text = "Verifica se o" & Chr(10) & "Gerenciador Padrão" & Chr(10) & "Está ativo!"
    'Call CriaLogTEF("ChamaSolicitacaoADM: Mensagem 23: " & rtxt_mensagem.Text & " às: " & Time)
    Call CriaLogTEF(Date & " " & Time & " ChamaSolicitacaoADM: " & rtxt_mensagem.Text & " - Bandeira=" & gBandeira)
    rtxt_mensagem.Visible = True
    DoEvents
    'If Not CopiaRegistrada Then
    '    gTefResposta = False
    '    Exit Sub
    'End If
    If UCase(gBandeira) = "TCSMART" Then
        gTefResposta = True
    Else
        gTefResposta = GerenciadorPadraoAtivo(False, True)
    End If
    rtxt_mensagem.Visible = False
    If gTefResposta Then
        'gTefString = "SolicitacaoADM"
        'ElseIf gTefString = "SolicitacaoADM" Then
        rtxt_mensagem.Text = "Solicitação de Funções" & Chr(10) & "Administrativas!"
        'Call CriaLogTEF("ChamaSolicitacaoADM: Mensagem 24: " & rtxt_mensagem.Text & " às: " & Time)
        rtxt_mensagem.Visible = True
        DoEvents
        gTefResposta = SolicitacaoADM
        Call CriaLogTefEspecial("ChamaSolicitacaoADM - 001 Retorno da rotina SolicitacaoADM. gTefResposta: " & gTefResposta)
        rtxt_mensagem.Visible = False
        If gTefResposta Then
            'gTefString = "TestaSolicitacaoADM"
            'ElseIf gTefString = "TestaSolicitacaoADM" Then
            gTefResposta = TestaSolicitacaoADM
            Call CriaLogTefEspecial("ChamaSolicitacaoADM - 002 Retorno da rotina TestaSolicitacaoADM. gTefResposta: " & gTefResposta)
            If gTefResposta Then
                'gTefString = "ImprimeTefADM"
                'ElseIf gTefString = "ImprimeTefADM" Then
                '''rtxt_mensagem.Text = "Aguarde!" & Chr(10) & Chr(10) & "Imprimindo TEF" & Chr(10) & "Administrativo."
                'Call CriaLogTEF("ChamaSolicitacaoADM: Mensagem 25: " & rtxt_mensagem.Text & " às: " & Time)
                rtxt_mensagem.Visible = True
                DoEvents
                gTefResposta = TestaImprimeTefADM("LeituraX")
                Call CriaLogTefEspecial("ChamaSolicitacaoADM - 003 Retorno da rotina TestaImprimeTefADM. gTefResposta: " & gTefResposta)
                'rtxt_mensagem.Visible = False
                ''''SolicitacaoADM = True
                If gTefResposta Then
                    'Aguarda 3 segundos para acabar de imprimir
                    'o tef, e depois sair a mensagem de confirmação
                    xHoraInicial = Time
                    Do Until DateDiff("s", xHoraInicial, Time) >= 3
                        DoEvents
                    Loop
                    'MsgBox "Confirma CNF"
                    'gTefString = "CNF"
                    'ElseIf gTefString = "CNF" Then
                    rtxt_mensagem.Text = "Aguarde!" & Chr(10) & Chr(10) & "Confirmando TEF."
                    'Call CriaLogTEF("ChamaSolicitacaoADM: Mensagem 26: " & rtxt_mensagem.Text & " às: " & Time)
                    rtxt_mensagem.Visible = True
                    DoEvents
                    gTefResposta = CNF
                    Call CriaLogTefEspecial("ChamaSolicitacaoADM - 004 Retorno da rotina CNF. gTefResposta: " & gTefResposta)
                    rtxt_mensagem.Visible = False
                Else
                    'MsgBox "Cancela NCN"
                    'gTefString = "NCN"
                    'ElseIf gTefString = "NCN" Then
                    rtxt_mensagem.Text = "Aguarde!" & Chr(10) & Chr(10) & "Cancelando TEF."
                    'Call CriaLogTEF("ChamaSolicitacaoADM: Mensagem 27: " & rtxt_mensagem.Text & " às: " & Time)
                    rtxt_mensagem.Visible = True
                    DoEvents
                    gTefResposta = SolicitacaoNCN
                    Call CriaLogTefEspecial("ChamaSolicitacaoADM - 005 Retorno da rotina SolicitacaoNCN. gTefResposta: " & gTefResposta)
                    rtxt_mensagem.Visible = False
                End If
            Else
                'MsgBox "SOLICITACAO ADMINISTRATIVA NÃO EFETUADO!", vbInformation, "ADM 1"
            End If
        'Else
        '    MsgBox "SOLICITACAO ADMINISTRATIVA NÃO EFETUADO!", vbInformation, "ADM 2"
        End If
    End If

End Sub
Private Sub ChamaSolicitacaoAlteraPrecoTCS()
    Dim xHoraInicial As Date
    Dim retval As Long
    
    rtxt_mensagem.Visible = True
    DoEvents
    'If Not CopiaRegistrada Then
    '    gTefResposta = False
    '    Exit Sub
    'End If
    retval = Shell("C:\TCS\tcs.exe /A", vbNormalFocus)
    'Aguarda 10 segundos
    xHoraInicial = Time
    Do Until DateDiff("s", xHoraInicial, Time) >= 10
        DoEvents
    Loop
    gTefResposta = TestaAlteraPrecoTCS
End Sub
Private Sub ChamaSolicitacaoCNC()
    rtxt_mensagem.Text = "Verifica se o" & Chr(10) & "Gerenciador Padrão" & Chr(10) & "Está ativo!"
    'Call CriaLogTEF("ChamaSolicitacaoCNC: Mensagem 28: " & rtxt_mensagem.Text & " às: " & Time)
    Call CriaLogTEF(Date & " " & Time & " ChamaSolicitacaoCNC: " & rtxt_mensagem.Text & " - Bandeira=" & gBandeira)
    rtxt_mensagem.Visible = True
    DoEvents
    gTefResposta = GerenciadorPadraoAtivo(False, True)
    rtxt_mensagem.Visible = False
    If gTefResposta Then
        '''rtxt_mensagem.Text = "Aguarde!" & Chr(10) & Chr(10) & "Solicitando Cancelamento" & Chr(10) & "de Venda TEF."
        'Call CriaLogTEF("ChamaSolicitacaoCNC: Mensagem 29: " & rtxt_mensagem.Text & " às: " & Time)
        rtxt_mensagem.Visible = True
        DoEvents
        gTefResposta = SolicitacaoCNC
        rtxt_mensagem.Visible = False
        If gTefResposta Then
            gTefResposta = TestaSolicitacao
            If gTefResposta Then
                'gTefString = "ImprimeTefADM"
                'ElseIf gTefString = "ImprimeTefADM" Then
                '''rtxt_mensagem.Text = "Aguarde!" & Chr(10) & Chr(10) & "Imprimindo TEF" & Chr(10) & "CNC."
                'Call CriaLogTEF("ChamaSolicitacaoCNC: Mensagem 30: " & rtxt_mensagem.Text & " às: " & Time)
                rtxt_mensagem.Visible = True
                DoEvents
                gTefResposta = TestaImprimeTefADM("LeituraX")
                rtxt_mensagem.Visible = False
                ''''SolicitacaoADM = True
                If gTefResposta Then
                    'MsgBox "Confirma CNF"
                    'gTefString = "CNF"
                    'ElseIf gTefString = "CNF" Then
                    rtxt_mensagem.Text = "Aguarde!" & Chr(10) & Chr(10) & "Confirmando TEF."
                    'Call CriaLogTEF("ChamaSolicitacaoCNC: Mensagem 31: " & rtxt_mensagem.Text & " às: " & Time)
                    rtxt_mensagem.Visible = True
                    DoEvents
                    gTefResposta = CNF
                    rtxt_mensagem.Visible = False
                Else
                    'MsgBox "Cancela NCN"
                    'gTefString = "NCN"
                    'ElseIf gTefString = "NCN" Then
                    rtxt_mensagem.Text = "Aguarde!" & Chr(10) & Chr(10) & "Cancelando TEF."
                    'Call CriaLogTEF("ChamaSolicitacaoCNC: Mensagem 32: " & rtxt_mensagem.Text & " às: " & Time)
                    rtxt_mensagem.Visible = True
                    DoEvents
                    gTefResposta = SolicitacaoNCN
                    rtxt_mensagem.Visible = False
                End If
            ''Else
            ''    MsgBox "Selecione outra forma de pagamento!!", vbInformation, "Forma de Pagamento Temporariamente Não Aceita!"
            ''    cbo_forma_pagamento.SetFocus
            ''    Exit Sub
            End If
        ''Else
        ''    MsgBox "SOLICITACAO DE COMPRA NAO EFETUADO"
        End If
    End If
End Sub
Private Sub ChamaSolicitacaoConsultaCH()
    Dim xHoraInicial As Date
    
    rtxt_mensagem.Text = "Verifica se o" & Chr(10) & "Gerenciador Padrão" & Chr(10) & "Está ativo!"
    'Call CriaLogTEF("ChamaSolicitacaoConsultaCH: Mensagem 23: " & rtxt_mensagem.Text & " às: " & Time)
    Call CriaLogTEF(Date & " " & Time & " ChamaSolicitacaoConsultaCH: " & rtxt_mensagem.Text & " - Bandeira=" & gBandeira)
    rtxt_mensagem.Visible = True
    DoEvents
    'If Not CopiaRegistrada Then
    '    gTefResposta = False
    '    Exit Sub
    'End If
    If UCase(gBandeira) = "TCSMART" Then
        gTefResposta = True
    Else
        gTefResposta = GerenciadorPadraoAtivo(False, True)
    End If
    rtxt_mensagem.Visible = False
    If gTefResposta Then
        rtxt_mensagem.Text = "Solicitação de Funções" & Chr(10) & "Consulta de Cheque Serasa!"
        'Call CriaLogTEF("ChamaSolicitacaoConsultaCH: Mensagem 24: " & rtxt_mensagem.Text & " às: " & Time)
        rtxt_mensagem.Visible = True
        DoEvents
        gTefResposta = SolicitacaoConsultaCH
        rtxt_mensagem.Visible = False
        If gTefResposta Then
            gTefResposta = TestaSolicitacaoADM
            If gTefResposta Then
                rtxt_mensagem.Visible = True
                DoEvents
                gTefResposta = TestaImprimeTefADM("LeituraX")
                'rtxt_mensagem.Visible = False
                ''''SolicitacaoADM = True
                If gTefResposta Then
                    'Aguarda 3 segundos para acabar de imprimir
                    'o tef, e depois sair a mensagem de confirmação
                    xHoraInicial = Time
                    Do Until DateDiff("s", xHoraInicial, Time) >= 3
                        DoEvents
                    Loop
                    rtxt_mensagem.Text = "Aguarde!" & Chr(10) & Chr(10) & "Confirmando TEF."
                    'Call CriaLogTEF("ChamaSolicitacaoConsultaCH: Mensagem 26: " & rtxt_mensagem.Text & " às: " & Time)
                    rtxt_mensagem.Visible = True
                    DoEvents
                    gTefResposta = CNF
                    rtxt_mensagem.Visible = False
                Else
                    rtxt_mensagem.Text = "Aguarde!" & Chr(10) & Chr(10) & "Cancelando TEF."
                    'Call CriaLogTEF("ChamaSolicitacaoConsultaCH: Mensagem 27: " & rtxt_mensagem.Text & " às: " & Time)
                    rtxt_mensagem.Visible = True
                    DoEvents
                    gTefResposta = SolicitacaoNCN
                    rtxt_mensagem.Visible = False
                End If
            Else
                'MsgBox "SOLICITACAO CONSULTA CHEQUE SERASA NÃO EFETUADO!", vbInformation, "CONSULTA CH 1"
            End If
        'Else
        '    MsgBox "SOLICITACAO CONSULTA CHEQUE SERASA NÃO EFETUADO!", vbInformation, "CONSULTA CH 2"
        End If
    End If

End Sub
Private Sub ChamaSolicitacaoNCN()
    rtxt_mensagem.Text = "Verifica se o" & Chr(10) & "Gerenciador Padrão" & Chr(10) & "Está ativo!"
    Call CriaLogTEF(Date & " " & Time & " ChamaSolicitacaoNCN: " & rtxt_mensagem.Text & " - Bandeira=" & gBandeira)
    rtxt_mensagem.Visible = True
    DoEvents
    If UCase(gBandeira) = "TCSMART" Then
        gTefResposta = True
    Else
        gTefResposta = LoopGerenciadorPadrao
        'gTefResposta = GerenciadorPadraoAtivo(False, True)
    End If
    'rtxt_mensagem.Visible = False
    If gTefResposta Then
        rtxt_mensagem.Text = "Aguarde!" & Chr(10) & Chr(10) & "Resolvendo Pendência."
        'Call CriaLogTEF("ChamaSolicitacaoNCN: Mensagem 34: " & rtxt_mensagem.Text & " às: " & Time)
        rtxt_mensagem.Visible = True
        DoEvents
        gTefResposta = TestaPendencia
        'rtxt_mensagem.Visible = False
        If gTefResposta Then
            gTefResposta = SolicitacaoNCN
            'If gTefResposta Then
            '    MsgBox "sim"
            'Else
            '    MsgBox "nao"
            'End If
        End If
    End If
End Sub
Private Sub ChamaSolicitacaoDSC()
    Dim xHoraInicial As Date
    
    'gTefString = "GerenciadorPadraoAtivo"
    'If gTefString = "GerenciadorPadraoAtivo" Then
    rtxt_mensagem.Text = "Verifica se o" & Chr(10) & "Gerenciador Padrão" & Chr(10) & "Está ativo!"
    Call CriaLogTEF(Date & " " & Time & " ChamaSolicitacaoDSC: " & rtxt_mensagem.Text & " - ECF N." & gNumeroCupom & " - Valor=" & gValorRecebido)
    rtxt_mensagem.Visible = True
    DoEvents
    'If Not CopiaRegistrada Then
    '    gTefResposta = False
    '    Exit Sub
    'End If
    gTefResposta = GerenciadorPadraoAtivo(False, True)
    rtxt_mensagem.Visible = False
    If gTefResposta Then
        'gTefString = "SolicitacaoDeDesconto"
        'ElseIf gTefString = "SolicitacaoDeDesconto" Then
        rtxt_mensagem.Text = "Solicitação de Desconto!"
        'Call CriaLogTEF("ChamaSolicitacaoDSC: Mensagem 2: " & rtxt_mensagem.Text & " às: " & Time)
        rtxt_mensagem.Visible = True
        DoEvents
        gTefResposta = SolicitacaoDeDesconto
        rtxt_mensagem.Visible = False
        If gTefResposta Then
            'gTefString = "TestaSolicitacaoDeDesconto"
            'ElseIf gTefString = "TestaSolicitacaoDeDesconto" Then
            gTefResposta = TestaSolicitacao
            If gTefResposta Then
                '''SolicitacaoDSC = True
                'gTefString = "ImprimeDSC"
                'ElseIf gTefString = "ImprimeDSC" Then
                '''rtxt_mensagem.Text = "Aguarde!" & Chr(10) & Chr(10) & "Imprimindo Desconto."
                'Call CriaLogTEF("ChamaSolicitacaoDSC: Mensagem 3: " & rtxt_mensagem.Text & " às: " & Time)
                rtxt_mensagem.Visible = True
                DoEvents
                'MsgBox "Confirma CNF"
                'gTefString = "CNF"
                'ElseIf gTefString = "CNF" Then
                rtxt_mensagem.Text = "Aguarde!" & Chr(10) & Chr(10) & "Confirmando Desconto."
                'Call CriaLogTEF("ChamaSolicitacaoDSC: Mensagem 4: " & rtxt_mensagem.Text & " às: " & Time)
                rtxt_mensagem.Visible = True
                DoEvents
                'Linha abaixo é para imprimir Desconto Concedido sem testar nada
                If gTipoDocumentoFiscal = "NFCe" Then
                    gTefResposta = TestaImprimeTEF("Vinculado")
                End If
                gTefResposta = CNF
                Call CriaLogTEF(Date & " " & Time & " ChamaSolicitacaoDSC: CNF voltou da confirmação com gTefResposta=" & gTefResposta)
                If gTefResposta = True Then
'                    gString = "Desconto Concluído!" & "|@|"
'                    gString = gString & "Desconto Confirmado com sucesso!" & "|@|"
'                    gString = gString & "30" & "|@|"
'                    frmMensagemAutomatica.Show 1
                Else
                    MsgBox "Erro na confirmação automática do Desconto." & Chr(10) & "Este desconto não será registrado ao estabelecimento!" & Chr(10) & "Este desconto não será calculado para o cliente!" & Chr(10) & "Para conceder o desconto, passe o cartão novamente!", vbCritical, "Erro: Desconto não foi Confirmado!"
                End If
                rtxt_mensagem.Visible = False
            End If
        End If
    End If
   
End Sub
Private Sub ChamaSolicitacaoTEF()
    Dim xHoraInicial As Date
    
    'gTefString = "GerenciadorPadraoAtivo"
    'If gTefString = "GerenciadorPadraoAtivo" Then
    rtxt_mensagem.Text = "Verifica se o" & Chr(10) & "Gerenciador Padrão" & Chr(10) & "Está ativo!"
    Call CriaLogTEF(Date & " " & Time & " ChamaSolicitacaoTEF: " & rtxt_mensagem.Text & " - ECF N." & gNumeroCupom & " - Valor=" & gValorRecebido)
    rtxt_mensagem.Visible = True
    DoEvents
    'If Not CopiaRegistrada Then
    '    gTefResposta = False
    '    Exit Sub
    'End If
    If UCase(gBandeira) = "TCSMART" Then
        gTefResposta = True
    Else
        gTefResposta = GerenciadorPadraoAtivo(False, True)
    End If
    rtxt_mensagem.Visible = False
    If gTefResposta Then
        'gTefString = "SolicitacaoDeCompra"
        'ElseIf gTefString = "SolicitacaoDeCompra" Then
        rtxt_mensagem.Text = "Solicitação de Compra!"
        'Call CriaLogTEF("ChamaSolicitacaoTEF: Mensagem 2: " & rtxt_mensagem.Text & " às: " & Time)
        rtxt_mensagem.Visible = True
        DoEvents
        gTefResposta = SolicitacaoDeCompra
        rtxt_mensagem.Visible = False
        If gTefResposta Then
            'gTefString = "TestaSolicitacaoDeCompra"
            'ElseIf gTefString = "TestaSolicitacaoDeCompra" Then
            gTefResposta = TestaSolicitacao
            If gTefResposta Then
                '''SolicitacaoTEF = True
                If ImprimeEncerramentoCupomFiscal Then
                    'gTefString = "ImprimeTEF"
                    'ElseIf gTefString = "ImprimeTEF" Then
                    If gConsultaCheque = False Or (gConsultaCheque = True And lCampo028 > 0) Then
                        '''rtxt_mensagem.Text = "Aguarde!" & Chr(10) & Chr(10) & "Imprimindo TEF."
                        'Call CriaLogTEF("ChamaSolicitacaoTEF: Mensagem 3: " & rtxt_mensagem.Text & " às: " & Time)
                        rtxt_mensagem.Visible = True
                        DoEvents
                        gTefResposta = TestaImprimeTEF("Vinculado")
                        'rtxt_mensagem.Visible = False
                        If gTefResposta Then
                            'Aguarda 3 segundos para acabar de imprimir
                            'o tef, e depois sair a mensagem de confirmação
                            xHoraInicial = Time
                            Do Until DateDiff("s", xHoraInicial, Time) >= 3
                                DoEvents
                            Loop
                            'MsgBox "Confirma CNF"
                            'gTefString = "CNF"
                            'ElseIf gTefString = "CNF" Then
                            rtxt_mensagem.Text = "Aguarde!" & Chr(10) & Chr(10) & "Confirmando TEF."
                            'Call CriaLogTEF("ChamaSolicitacaoTEF: Mensagem 4: " & rtxt_mensagem.Text & " às: " & Time)
                            rtxt_mensagem.Visible = True
                            DoEvents
                            gTefResposta = CNF
                            Call CriaLogTEF(Date & " " & Time & " ChamaSolicitacaoTEF: CNF voltou da confirmação com gTefResposta=" & gTefResposta)
                            If gTefResposta = True Then
                                gString = "TEF Concluído!" & "|@|"
                                gString = gString & "TEF Confirmado com sucesso!" & "|@|"
                                gString = gString & "3" & "|@|"
                                frmMensagemAutomatica.Show 1
                            Else
                                MsgBox "Erro na confirmação automática do TEF." & Chr(10) & "Este valor não será creditado ao estabelecimento!" & Chr(10) & "Este valor não será debitado do cliente!" & Chr(10) & "Para nao faltar no caixa passe o cartão novamente!", vbCritical, "Erro: TEF não foi Confirmado!"
                            End If
                            rtxt_mensagem.Visible = False
                        Else
                            'MsgBox "Cancela NCN"
                            'gTefString = "NCN"
                            'ElseIf gTefString = "NCN" Then
                            rtxt_mensagem.Text = "Aguarde!" & Chr(10) & Chr(10) & "Cancelando TEF."
                            'Call CriaLogTEF("ChamaSolicitacaoTEF: Mensagem 5: " & rtxt_mensagem.Text & " às: " & Time)
                            rtxt_mensagem.Visible = True
                            DoEvents
                            gTefResposta = SolicitacaoNCN
                            rtxt_mensagem.Visible = False
                        End If
                    Else
                        rtxt_mensagem.Text = "Aguarde!" & Chr(10) & Chr(10) & "Confirmando TEF."
                        'Call CriaLogTEF("ChamaSolicitacaoTEF: Mensagem 4: " & rtxt_mensagem.Text & " às: " & Time)
                        rtxt_mensagem.Visible = True
                        DoEvents
                        gTefResposta = CNF
                        If gTefResposta = True Then
                            MsgBox "TEF Confirmado com sucesso!", vbInformation, "TEF Concluído"
                        Else
                            MsgBox "Erro na confirmação automática do TEF." & Chr(10) & "Este valor não será creditado ao estabelecimento!" & Chr(10) & "Este valor não será debitado do cliente!" & Chr(10) & "Para nao faltar no caixa passe o cartão novamente!", vbCritical, "Erro: TEF não foi Confirmado!"
                        End If
                        rtxt_mensagem.Visible = False
                    End If
                Else
                    rtxt_mensagem.Text = "Aguarde!" & Chr(10) & Chr(10) & "Cancelando TEF."
                    'Call CriaLogTEF("ChamaSolicitacaoTEF: Mensagem 5: " & rtxt_mensagem.Text & " às: " & Time)
                    rtxt_mensagem.Visible = True
                    DoEvents
                    gTefResposta = SolicitacaoNCN
                    rtxt_mensagem.Visible = False
                End If
            End If
        End If
    End If
   
End Sub
Function SolicitacaoCNC() As Boolean
    Dim xString As String
    Dim xValorRecebido As String
    Dim xNomeRede As String
    Dim xNumeroTransacao As String
    Dim xNomeArquivo As String     'Arquivo a Ser Criado "ATV"
    Dim xNomeArquivoResp As String 'Arquivo de Resposta
    Dim xNomeArquivoTemp As String 'Arquivo Temporário
    
    On Error GoTo FileError
    
    SolicitacaoCNC = False
'    If UCase(gBandeira) = "TECBAN" Then
'        xNomeArquivo = "C:\TEF_DISC\REQ\IntPos.001"
'        xNomeArquivoResp = "C:\TEF_DISC\RESP\IntPos.STS"
'        xNomeArquivoTemp = "C:\TEF_DISC\REQ\IntPos.tmp"
'    ElseIf UCase(gBandeira) = "SMARTEF" Then
'        xNomeArquivo = "C:\SMARTEF\REQ\IntPos.001"
'        xNomeArquivoResp = "C:\SMARTEF\RESP\IntPos.STS"
'        xNomeArquivoTemp = "C:\SMARTEF\REQ\IntPos.tmp"
'    ElseIf UCase(gBandeira) = "SUPERTEF" Then
'        xNomeArquivo = "C:\SUPERTEF\REQ\IntPos.001"
'        xNomeArquivoResp = "C:\SUPERTEF\RESP\IntPos.STS"
'        xNomeArquivoTemp = "C:\SUPERTEF\REQ\IntPos.tmp"
'    ElseIf UCase(gBandeira) = "HIPERTEF" Then
'        xNomeArquivo = "C:\HiperTEF\REQ\IntPos.001"
'        xNomeArquivoResp = "C:\HiperTEF\RESP\IntPos.STS"
'        xNomeArquivoTemp = "C:\HiperTEF\REQ\IntPos.tmp"
'    ElseIf UCase(gBandeira) = "PAGCARD" Then
'        xNomeArquivo = "C:\CardTech_NEUS\REQ\IntPos.001"
'        xNomeArquivoResp = "C:\CardTech_NEUS\RESP\IntPos.STS"
'        xNomeArquivoTemp = "C:\CardTech_NEUS\REQ\IntPos.tmp"
'    Else
'        xNomeArquivo = "C:\TEF_DIAL\REQ\IntPos.001"
'        xNomeArquivoResp = "C:\TEF_DIAL\RESP\IntPos.STS"
'        xNomeArquivoTemp = "C:\TEF_DIAL\REQ\IntPos.tmp"
'    End If
    xNomeArquivo = gDiretorioReq & "IntPos.001"
    xNomeArquivoResp = gDiretorioResp & "IntPos.STS"
    xNomeArquivoTemp = gDiretorioReq & "IntPos.tmp"
    
    'Se existir arquivos, deleta.
    If gArqTxt.FileExists(xNomeArquivoTemp) Then
        Call ExcluirArquivo(xNomeArquivoTemp)
    End If
    If gArqTxt.FileExists(xNomeArquivo) Then
        Call ExcluirArquivo(xNomeArquivo)
    End If
    
    'CNC
    Set gArquivo = gArqTxt.CreateTextFile(xNomeArquivoTemp)
    xNomeRede = UCase(gBandeira)
    gArquivo.WriteLine ("000-000 = CNC")
    gArquivo.WriteLine ("001-000 = " & Format(gNumeroControleSolicitacao, "0000000000"))
    '''gArquivo.WriteLine ("003-000 = " & Mid(Format(fValidaValor(txt_valor.Text), "0000000000.00"), 1, 10) & Mid(Format(fValidaValor(txt_valor.Text), "0000000000.00"), 12, 2)) 'VALOR
    gArquivo.WriteLine ("003-000 = " & fValidaValor(txt_valor.Text) * 100)
    gArquivo.WriteLine ("010-000 = " & xNomeRede)
    gArquivo.WriteLine ("012-000 = " & txt_nsu.Text)
    gArquivo.WriteLine ("022-000 = " & Format(Date, "ddmmyyyy")) 'txt_data.Text)  'DATA
    gArquivo.WriteLine ("023-000 = " & Format(Time, "hhmmss"))   'txt_hora.Text)  'HORA
    If gTipoDesconto = "POSTOAKI" Then
        gArquivo.WriteLine ("600-000 = " & gCodigoColaborador)
        gArquivo.WriteLine ("601-000 = " & gNomeColaborador)
        gArquivo.WriteLine ("602-000 = " & gAvaliacaoColaborador)
    End If
    gArquivo.WriteLine ("999-999 = 0")
    gArquivo.Close

    If gArqTxt.FileExists(xNomeArquivoTemp) Then
        Call RenomearArquivo(xNomeArquivoTemp, xNomeArquivo)
    End If
    
    'Aguarda 7 segundos
    lHoraInicial = Time
    Do Until DateDiff("s", lHoraInicial, Time) >= 7
        If gArqTxt.FileExists(xNomeArquivoResp) Then
            'BUG:XP DÁ UM TEMPO PARA LIBERAR ARQUIVO PARA LEITURA
            lHoraInicial = Time
            Do Until DateDiff("s", lHoraInicial, Time) >= 2
            Loop
            Exit Do
        End If
        DoEvents
    Loop
    
    If gArqTxt.FileExists(xNomeArquivoResp) Then
        Set gArquivo = gArqTxt.OpenTextFile(xNomeArquivoResp, ForReading)
        lString = gArquivo.ReadLine
        If lString = "000-000 = CNC" Then
            Do Until gArquivo.AtEndOfStream
                lString = gArquivo.ReadLine
                'Verifica se o número do controle da solicitação é igual
                If Mid(lString, 1, 7) = "001-000" Then
                    If Mid(lString, 11, 10) = gNumeroControleSolicitacao Then
                        lCampo001 = True
                    End If
                End If
            Loop
        End If
        gArquivo.Close
        Call ExcluirArquivo(xNomeArquivoResp)
        If lCampo001 Then
            SolicitacaoCNC = True
            Call CriaLogTEF(Date & " " & Time & " Transação TEF Cancelada: CNC " & rtxt_mensagem.Text & " - Bandeira=" & xNomeRede & " - Valor=" & txt_valor.Text & " - Numero da Transação=" & gNumeroControleSolicitacao)
        Else
            MsgBox "Numero de controle de solicitação está diferente", vbInformation, "SolicitacaoCNC"
        End If
    Else
        'Como não teve respostas em 7 segundos
        'Deleta Arquivo de Requisicao
        If gArqTxt.FileExists(xNomeArquivo) Then
            Call ExcluirArquivo(xNomeArquivo)
        End If
        MsgBox "TEF Não Está Ativo!", vbInformation, "Mensagem Padrão"
    End If
    Exit Function
FileError:
    Call CriaLogTEF(Date & " " & Time & " Erro SolicitacaoCNC: " & " - ErroNúmero: " & Err & " - ErroTexto: " & Error)
    MsgBox "ERRO não identificado!", vbInformation, "Rotina: SolicitacaoCNC"
    Exit Function
End Function
Function SolicitacaoConsultaCH() As Boolean
    Dim i As Integer
    Dim xNomeArquivo As String        'Arquivo a Ser Criado "CHQ"
    Dim xNomeArquivoResp As String    'Arquivo de Resposta
    Dim xNomeArquivoResp001 As String 'Arquivo de Resposta
    Dim xNomeArquivoTemp As String    'Arquivo Temporário
    Dim xString As String
    Dim retval As Long
    
    On Error GoTo FileError
    
    SolicitacaoConsultaCH = False
    lCampo001 = False
    If UCase(gBandeira) = "TCSMART" Then
        xNomeArquivo = gDiretorioReq & "IntTcs.001"
        xNomeArquivoResp = gDiretorioResp & "IntTcs.RET"
        xNomeArquivoResp001 = gDiretorioResp & "IntTcs.001"
        xNomeArquivoTemp = gDiretorioReq & "IntTcs.tmp"
    Else
        xNomeArquivo = gDiretorioReq & "IntPos.001"
        xNomeArquivoResp = gDiretorioResp & "IntPos.STS"
        xNomeArquivoResp001 = gDiretorioResp & "IntPos.001"
        xNomeArquivoTemp = gDiretorioReq & "IntPos.tmp"
    End If
    
    'CHQ
    'Se existir arquivos, deleta.
    If gArqTxt.FileExists(xNomeArquivoTemp) Then
        Call ExcluirArquivo(xNomeArquivoTemp)
    End If
    If gArqTxt.FileExists(xNomeArquivo) Then
        Call ExcluirArquivo(xNomeArquivo)
    End If
    
    
    'Gera Arquivo
    Set gArquivo = gArqTxt.CreateTextFile(xNomeArquivoTemp)
    If UCase(gBandeira) = "TCSMART" Then
        gArquivo.WriteLine ("500-000 = CHQ")
        gArquivo.WriteLine ("501-000 = " & Format(gNumeroControleSolicitacao, "0000000000"))
    Else
        gArquivo.WriteLine ("000-000 = CHQ")
        gArquivo.WriteLine ("001-000 = " & Format(gNumeroControleSolicitacao, "0000000000"))
        gArquivo.WriteLine ("002-000 = " & Format(gNumeroCupom + 1, "000000"))
        lString = Format(fValidaValor(gValorRecebido), "#########0.00")
        i = Len(lString)
        lString = Mid(lString, 1, i - 3) & Mid(lString, i - 1, 2)
        gArquivo.WriteLine ("003-000 = " & lString)
    End If
    gArquivo.WriteLine ("999-999 = 0")
    gArquivo.Close

    
    If gArqTxt.FileExists(xNomeArquivoTemp) Then
        Call RenomearArquivo(xNomeArquivoTemp, xNomeArquivo)
        End If
    
    If UCase(gBandeira) = "TCSMART" Then
        retval = Shell("C:\TCS\tcs.exe", vbNormalFocus)
        Do Until gArqTxt.FileExists(xNomeArquivoResp)
            DoEvents
        Loop
    End If
    'Aguarda 7 segundos
    lHoraInicial = Time
    Do Until DateDiff("s", lHoraInicial, Time) >= 7
        If gArqTxt.FileExists(xNomeArquivoResp) Then
            If UCase(gBandeira) <> "TEFCERRADO" Then
                'BUG:XP DÁ UM TEMPO PARA LIBERAR ARQUIVO PARA LEITURA
                lHoraInicial = Time
                Do Until DateDiff("s", lHoraInicial, Time) >= 2
                Loop
            End If
            Exit Do
        End If
        DoEvents
    Loop
    
    If Not gArqTxt.FileExists(xNomeArquivoResp) Then
        MsgBox "Gerenciador Padrão não está ativo, e será ativado automaticamente!", vbInformation, "Mensagem Padrão"
        Call AtivaGerenciadorPadrao
        'Aguarda 7 segundos
        lHoraInicial = Time
        Do Until DateDiff("s", lHoraInicial, Time) >= 7
            If gArqTxt.FileExists(xNomeArquivoResp) Then
                If UCase(gBandeira) <> "TEFCERRADO" Then
                    'BUG:XP DÁ UM TEMPO PARA LIBERAR ARQUIVO PARA LEITURA
                    lHoraInicial = Time
                    Do Until DateDiff("s", lHoraInicial, Time) >= 2
                    Loop
                End If
                Exit Do
            End If
            DoEvents
        Loop
    End If
    
    If gArqTxt.FileExists(xNomeArquivoResp) Then
        Set gArquivo = gArqTxt.OpenTextFile(xNomeArquivoResp, ForReading)
        lString = gArquivo.ReadLine
        If UCase(gBandeira) = "TCSMART" Then
            If lString = "517-000 = 000" Then
                gArquivo.Close
                Set gArquivo = gArqTxt.OpenTextFile(xNomeArquivoResp001, ForReading)
                Do Until gArquivo.AtEndOfStream
                    lString = gArquivo.ReadLine
                    'Verifica se o número do controle da solicitação é igual
                    If Mid(lString, 1, 7) = "501-000" Then
                        If Mid(lString, 11, 10) = gNumeroControleSolicitacao Then
                            lCampo001 = True
                            Exit Do
                        End If
                    End If
                Loop
            Else
                lString = gArquivo.ReadLine
                lString = Mid(lString, 12, Len(lString) - 12)
            End If
        Else
            If lString = "000-000 = CHQ" Then
                Do Until gArquivo.AtEndOfStream
                    lString = gArquivo.ReadLine
                    'Verifica se o número do controle da solicitação é igual
                    If Mid(lString, 1, 7) = "001-000" Then
                        If Mid(lString, 11, 10) = gNumeroControleSolicitacao Then
                            lCampo001 = True
                            Exit Do
                        End If
                    End If
                Loop
            End If
        End If
        gArquivo.Close
        
        Call ExcluirArquivo(xNomeArquivoResp)
        If lCampo001 Then
            SolicitacaoConsultaCH = True
        Else
            If UCase(gBandeira) = "TCSMART" Then
                MsgBox lString, vbInformation, "Solicitação Consulta Cheque Serasa"
            Else
                MsgBox "Numero de controle de solicitação está diferente", vbInformation, "SolicitacaoConsultaCH"
            End If
        End If
    Else
        'Como não teve respostas em 7 segundos
        'Deleta Arquivo de Requisicao
        If gArqTxt.FileExists(xNomeArquivo) Then
            Call ExcluirArquivo(xNomeArquivo)
        End If
        MsgBox "Gerenciador Padrão não está ativo!", vbInformation, "Mensagem Padrão"
    End If
    Exit Function
FileError:
    Call CriaLogTEF(Date & " " & Time & " SolicitacaoConsultaCH: " & " - ErroNúmero: " & Err & " - ErroTexto: " & Error)
    If Err = 76 Then
        MsgBox "Gerenciador Padrão não está instalado neste computador!", vbInformation, "Transação com TEF não aceita!"
        Exit Function
    End If
    MsgBox "ERRO não identificado!" & Chr(10) & Error, vbInformation, "Rotina: SolicitacaoConsultaCH"
    Exit Function
End Function
Function CNF() As Boolean
    Dim xString As String
    Dim xValorRecebido As String
    Dim xNomeRede As String
    Dim xNumeroTransacao As String
    Dim xFinalizacao As String
    Dim i As Integer
    Dim i2 As Integer
    Dim xNomeArquivo As String        'Arquivo a Ser Criado "ATV"
    Dim xNomeArquivoResp As String    'Arquivo de Resposta
    Dim xNomeArquivoResp001 As String 'Arquivo de Resposta
    Dim xNomeArquivoTemp As String    'Arquivo Temporário
    
    Call CriaLogTEF(Date & " " & Time & " Iniciada CNF ")
    
    CNF = False

    If UCase(gBandeira) = "TCSMART" Then
        Call ExcluirArquivo(gDiretorioResp & "IntTCS.001")
        CNF = True
        Exit Function
    Else
        xNomeArquivo = gDiretorioReq & "IntPos.001"
        xNomeArquivoResp = gDiretorioResp & "IntPos.STS"
        xNomeArquivoResp001 = gDiretorioResp & "IntPos.001"
        xNomeArquivoTemp = gDiretorioReq & "IntPos.tmp"
    End If
    
    LoopGerenciadorPadrao
    
    'Deletar o backup do "IntPos.001" em c:\Backup_Sgp
    If gArqTxt.FileExists("C:\Backup_SGP\IntPos.001") Then
        Call ExcluirArquivo("C:\Backup_SGP\IntPos.001")
    End If
    
    If gArqTxt.FileExists(xNomeArquivoResp001) Then
        Call CriaLogTefEspecial("CNF - 001 Foi encontrado o arquivo xNomeArquivoResp001: " & xNomeArquivoResp001)
        Set gArquivo = gArqTxt.OpenTextFile(xNomeArquivoResp001, ForReading)
        lMensagem = ""
        i = 0
        lQtdMensagem29 = -1
        Do Until gArquivo.AtEndOfStream
            lString = gArquivo.ReadLine
            Call CriaLogTefEspecial("CNF - 002 Sequencia das linhas de retorno. lString: " & lString)
            'Guarda Texto da Finalizacao
            If Mid(lString, 1, 3) = "010" Then
                i2 = Len(lString)
                xNomeRede = Mid(lString, 11, i2 - 10)
            End If
            If Mid(lString, 1, 3) = "012" Then
                i2 = Len(lString)
                xNumeroTransacao = Mid(lString, 11, i2 - 10)
            End If
            If Mid(lString, 1, 3) = "027" Then
                i2 = Len(lString)
                xFinalizacao = Mid(lString, 11, i2 - 10)
                Exit Do
            End If
        Loop
        gArquivo.Close
    Else
        Call CriaLogTefEspecial("CNF - 003 Arquivo inexistente. xNomeArquivoResp001: " & xNomeArquivoResp001)
        MsgBox "O arquivo IntPos.001 foi deletado de forma não identificada!", vbInformation, "Mensagem Padrão"
        Exit Function
    End If
    
    If UCase(gBandeira) = "SMARTEF" Then
        Call ExcluirArquivo(xNomeArquivoResp001)
    End If
    
    
    Call CriaLogTefEspecial("CNF - 004 Solicitação anterior. gNumeroControleSolicitacao: " & gNumeroControleSolicitacao & " - gTipoDesconto: " & gTipoDesconto & " - gTefString: " & gTefString)
    gNumeroControleSolicitacao = lControleSolicitacao.ProximaSolicitacaoTEF(1)
    Call CriaLogTefEspecial("CNF - 005 Nova Solicitação. gNumeroControleSolicitacao: " & gNumeroControleSolicitacao)
    Set gArquivo = gArqTxt.CreateTextFile(xNomeArquivoTemp)
    gArquivo.WriteLine ("000-000 = CNF")
    gArquivo.WriteLine ("001-000 = " & Format(gNumeroControleSolicitacao, "0000000000"))
    gArquivo.WriteLine ("010-000 = " & xNomeRede)
    gArquivo.WriteLine ("012-000 = " & xNumeroTransacao)
    gArquivo.WriteLine ("027-000 = " & xFinalizacao)
    Call CriaLogTefEspecial("CNF - 006 Registros 000 e 001 gerados. gTefString=" & gTefString & " - xNomeRede=" & xNomeRede)
    If gTefString <> "SolicitacaoADM" Then
        Call CriaLogTefEspecial("CNF - 007 Registros 010, 012 e 027 gerados")
        If gTipoDesconto = "POSTOAKI" Then
            gArquivo.WriteLine ("600-000 = " & gCodigoColaborador)
            gArquivo.WriteLine ("601-000 = " & gNomeColaborador)
            gArquivo.WriteLine ("602-000 = " & gAvaliacaoColaborador)
            Call CriaLogTefEspecial("CNF - 008 Registros 600 a 602 gerados")
        End If
    End If
    gArquivo.WriteLine ("999-999 = 0")
    Call CriaLogTefEspecial("CNF - 009 Registros 999-999 gerado")
    gArquivo.Close
    
    If gArqTxt.FileExists(xNomeArquivoTemp) Then
        Call RenomearArquivo(xNomeArquivoTemp, xNomeArquivo)
    End If
    
    'Aguarda 7 segundos
    lHoraInicial = Time
    Do Until DateDiff("s", lHoraInicial, Time) >= 7
        If gArqTxt.FileExists(xNomeArquivoResp) Then
            'BUG:XP DÁ UM TEMPO PARA LIBERAR ARQUIVO PARA LEITURA
            lHoraInicial = Time
            Do Until DateDiff("s", lHoraInicial, Time) >= 2
            Loop
            Exit Do
        End If
        DoEvents
    Loop
    
    If gArqTxt.FileExists(xNomeArquivoResp) Then
        Call CriaLogTefEspecial("CNF - 010 Arquivo de retorno encontrado. xNomeArquivoResp: " & xNomeArquivoResp)
        Set gArquivo = gArqTxt.OpenTextFile(xNomeArquivoResp, ForReading)
        lCampo001 = False
        lString = gArquivo.ReadLine
        Call CriaLogTefEspecial("CNF - 011 Primeira linha de retorno. lString: " & lString)
        If lString = "000-000 = CNF" Then
            Do Until gArquivo.AtEndOfStream
                lString = gArquivo.ReadLine
                Call CriaLogTefEspecial("CNF - 012 Sequencia    de   retorno. lString: " & lString)
                If Mid(lString, 1, 7) = "001-000" Then
                    Call CriaLogTefEspecial("CNF - 013 Registro 001-000. Esperando gNumeroControleSolicitacao: " & gNumeroControleSolicitacao)
                    If Mid(lString, 11, 10) = gNumeroControleSolicitacao Then
                        Call CriaLogTefEspecial("CNF - 014 Campo lCampo001 passa a ser verdadeiro.")
                        lCampo001 = True
                    End If
                End If
            Loop
        End If
        gArquivo.Close
        Call ExcluirArquivo(xNomeArquivoResp)
        If lCampo001 Then
            Call ExcluirArquivo(xNomeArquivoResp001)
            CNF = True
            Call CriaLogTEF(Date & " " & Time & " Concluída CNF " & " - Bandeira=" & xNomeRede & " - Valor=" & txt_valor.Text & " - Numero da Transação=" & gNumeroControleSolicitacao)
        Else
            MsgBox "Numero de controle de solicitação está diferente" & Chr(10) & "Transação nao foi confirmada e será cancelada automaticamente!", vbCritical, "Erro Grave! CNF"
            Call CriaLogTEF(Date & " " & Time & " Erro rotina: CNF: foi informado ao usuário " & "Numero de controle de solicitação está diferente")
        End If
    Else
        Call CriaLogTefEspecial("CNF - 015 Arquivo de retorno INEXISTENTE. xNomeArquivoResp: " & xNomeArquivoResp)
        MsgBox "TEF Não Está Ativo!", vbInformation, "Mensagem Padrão"
        Call CriaLogTEF(Date & " " & Time & " Erro rotina: CNF: arquivo de resposta nao encontrado:" & xNomeArquivoResp)
    End If
End Function
Function GerenciadorPadraoAtivo(ByVal pAtivaAutomatico As Boolean, ByVal pNovoNumeroSolicitacao As Boolean) As Boolean
    Dim xNomeArquivo As String     'Arquivo a Ser Criado "ATV"
    Dim xNomeArquivoResp As String 'Arquivo de Resposta
    Dim xNomeArquivoTemp As String 'Arquivo Temporário
    Dim retval As Long
    Dim xFaseErro As Integer
    Dim xNumeroControleSolicitacao As Long
    
    On Error GoTo FileError
    
    xFaseErro = 0
    GerenciadorPadraoAtivo = False
    lCampo001 = False
    If UCase(gBandeira) = "SUPERTEF" Then
        GerenciadorPadraoAtivo = True
        lCampo001 = True
        Exit Function
    End If
    xFaseErro = 1
'    If UCase(gBandeira) = "TECBAN" Then
'        xNomeArquivo = "C:\TEF_DISC\REQ\IntPos.001"
'        xNomeArquivoResp = "C:\TEF_DISC\RESP\IntPos.STS"
'        xNomeArquivoTemp = "C:\TEF_DISC\REQ\IntPos.tmp"
'    ElseIf UCase(gBandeira) = "TCSMART" Then
'        xNomeArquivo = "C:\TCS\TX\IntTCS.001"
'        xNomeArquivoResp = "C:\TCS\RX\IntTCS.RET"
'        xNomeArquivoTemp = "C:\TCS\TX\IntTCS.tmp"
'    ElseIf UCase(gBandeira) = "SMARTEF" Then
'        xNomeArquivo = "C:\SMARTEF\REQ\IntPos.001"
'        xNomeArquivoResp = "C:\SMARTEF\RESP\IntPos.STS"
'        xNomeArquivoTemp = "C:\SMARTEF\REQ\IntPos.tmp"
'    ElseIf UCase(gBandeira) = "SUPERTEF" Then
'        xNomeArquivo = "C:\SUPERTEF\REQ\IntPos.001"
'        xNomeArquivoResp = "C:\SUPERTEF\RESP\IntPos.STS"
'        xNomeArquivoTemp = "C:\SUPERTEF\REQ\IntPos.tmp"
'    ElseIf UCase(gBandeira) = "HIPERTEF" Then
'        xNomeArquivo = "C:\HiperTEF\REQ\IntPos.001"
'        xNomeArquivoResp = "C:\HiperTEF\RESP\IntPos.STS"
'        xNomeArquivoTemp = "C:\HiperTEF\REQ\IntPos.tmp"
'    ElseIf UCase(gBandeira) = "PAGCARD" Then
'        xNomeArquivo = "C:\CardTech_NEUS\REQ\IntPos.001"
'        xNomeArquivoResp = "C:\CardTech_NEUS\RESP\IntPos.STS"
'        xNomeArquivoTemp = "C:\CardTech_NEUS\REQ\IntPos.tmp"
'    Else
'        xNomeArquivo = "C:\TEF_DIAL\REQ\IntPos.001"
'        xNomeArquivoResp = "C:\TEF_DIAL\RESP\IntPos.STS"
'        xNomeArquivoTemp = "C:\TEF_DIAL\REQ\IntPos.tmp"
'    End If
    If UCase(gBandeira) = "TCSMART" Then
        xNomeArquivo = gDiretorioReq & "IntTCS.001"
        xNomeArquivoResp = gDiretorioResp & "IntTCS.RET"
        xNomeArquivoTemp = gDiretorioReq & "IntTCS.tmp"
    Else
        xNomeArquivo = gDiretorioReq & "IntPos.001"
        xNomeArquivoResp = gDiretorioResp & "IntPos.STS"
        xNomeArquivoTemp = gDiretorioReq & "IntPos.tmp"
    End If
    
    xFaseErro = 2
    
    'ATV
    'Se existir arquivos, deleta.
    xFaseErro = 3
    If gArqTxt.FileExists(xNomeArquivoTemp) Then
        xFaseErro = 4
        Call ExcluirArquivo(xNomeArquivoTemp)
    End If
    xFaseErro = 5
    If gArqTxt.FileExists(xNomeArquivo) Then
        xFaseErro = 6
        Call ExcluirArquivo(xNomeArquivo)
    End If
    'Gera Arquivo
    
    xFaseErro = 7
    Set gArquivo = gArqTxt.CreateTextFile(xNomeArquivoTemp)
    
    xNumeroControleSolicitacao = gNumeroControleSolicitacao
    If pNovoNumeroSolicitacao Then
        xNumeroControleSolicitacao = lControleSolicitacao.ProximaSolicitacaoTEF(1)
    End If
    
    xFaseErro = 8
    If UCase(gBandeira) = "TCSMART" Then
        gArquivo.WriteLine ("500-000 = ADM")
        gArquivo.WriteLine ("501-000=" & Format(xNumeroControleSolicitacao, "0000000000"))
    Else
        gArquivo.WriteLine ("000-000 = ATV")
        gArquivo.WriteLine ("001-000 = " & Format(xNumeroControleSolicitacao, "0000000000"))
    End If
    xFaseErro = 9
    gArquivo.WriteLine ("999-999 = 0")
    gArquivo.Close

    xFaseErro = 10
    If gArqTxt.FileExists(xNomeArquivoTemp) Then
        xFaseErro = 11
        Call RenomearArquivo(xNomeArquivoTemp, xNomeArquivo)
    End If
    xFaseErro = 12
    If UCase(gBandeira) = "TCSMART" Then
        xFaseErro = 13
        retval = Shell("C:\TCS\tcs.exe", vbNormalFocus)
    End If
    'Aguarda 7 segundos
    lHoraInicial = Time
    xFaseErro = 14
    Do Until DateDiff("s", lHoraInicial, Time) >= 7
        If gArqTxt.FileExists(xNomeArquivoResp) Then
            'BUG:XP DÁ UM TEMPO PARA LIBERAR ARQUIVO PARA LEITURA
            lHoraInicial = Time
            Do Until DateDiff("s", lHoraInicial, Time) >= 2
            Loop
            Exit Do
        End If
        DoEvents
    Loop
    xFaseErro = 15
    
    'If Not gArqTxt.FileExists(xNomeArquivoResp) Then
    '    MsgBox "Gerenciador Padrão não está ativo.", vbInformation, "Mensagem Padrão"
        'Call AtivaGerenciadorPadrao
        'Aguarda 7 segundos
        'lHoraInicial = Time
        'Do Until DateDiff("s", lHoraInicial, Time) >= 7
        '    If gArqTxt.FileExists(xNomeArquivoResp) Then
        '        Exit Do
        '    End If
        '    DoEvents
        'Loop
    'End If
    
    If gArqTxt.FileExists(xNomeArquivoResp) Then
        xFaseErro = 16
        Set gArquivo = gArqTxt.OpenTextFile(xNomeArquivoResp, ForReading)
        xFaseErro = 17
        lString = gArquivo.ReadLine
        If UCase(gBandeira) = "TCSMART" Then
            If lString = "517-000 = 002" Then
                'Do Until gArquivo.AtEndOfStream
                '    lString = gArquivo.ReadLine
                '    'Verifica se o número do controle da solicitação é igual
                '    If Mid(lString, 1, 7) = "517-001" Then
                        lCampo001 = True
                '        Exit Do
                '    End If
                'Loop
            End If
        Else
            If lString = "000-000 = ATV" Then
                Do Until gArquivo.AtEndOfStream
                    lString = gArquivo.ReadLine
                    'Verifica se o número do controle da solicitação é igual
                    If Mid(lString, 1, 7) = "001-000" Then
                        If Mid(lString, 11, 10) = xNumeroControleSolicitacao Then
                            lCampo001 = True
                            Exit Do
                        End If
                    End If
                Loop
            End If
        End If
        gArquivo.Close
        Call ExcluirArquivo(xNomeArquivoResp)
        'If UCase(gBandeira) = "TCSMART" Then
        '    xNomeArquivoResp = Replace(xNomeArquivoResp, "RET", "001")
        '    Call ExcluirArquivo(xNomeArquivoResp)
        'End If
        If lCampo001 Then
            GerenciadorPadraoAtivo = True
        Else
            MsgBox "Numero de controle de solicitação está diferente", vbInformation, "GerenciadorPadraoAtivo"
            Call ExcluirArquivo("C:\TEF_DISC\RESP\*.STS")
            Call ExcluirArquivo("C:\TEF_DISC\RESP\*.001")
            Call ExcluirArquivo("C:\TEF_DIAL\RESP\*.STS")
            Call ExcluirArquivo("C:\TEF_DIAL\RESP\*.001")
            Call ExcluirArquivo("C:\SUPERTEF\RESP\*.STS")
            Call ExcluirArquivo("C:\SUPERTEF\RESP\*.001")
            Call ExcluirArquivo("C:\HiperTEF\RESP\*.STS")
            Call ExcluirArquivo("C:\HiperTEF\RESP\*.001")
            Call ExcluirArquivo("C:\CardTech_NEUS\RESP\*.STS")
            Call ExcluirArquivo("C:\CardTech_NEUS\RESP\*.001")
            Call ExcluirArquivo("C:\TefCerrado\RESP\*.STS")
            Call ExcluirArquivo("C:\TefCerrado\RESP\*.001")
            Call ExcluirArquivo("C:\Tef_NEUS\RESP\*.STS")
            Call ExcluirArquivo("C:\Tef_NEUS\RESP\*.001")
            Call ExcluirArquivo("C:\GetNet\RESP\*.STS")
            Call ExcluirArquivo("C:\GetNet\RESP\*.001")
        End If
    Else
        'Como não teve respostas em 7 segundos
        'Deleta Arquivo de Requisicao
        If gArqTxt.FileExists(xNomeArquivo) Then
            Call ExcluirArquivo(xNomeArquivo)
        End If
        If pAtivaAutomatico Then
            'nao recomendado fazer agora, pois seria muita mudanca
            MsgBox "Gerenciador Padrão não está ativo!", vbInformation, "Mensagem Padrão"
        Else
            MsgBox "Gerenciador Padrão não está ativo!", vbInformation, "Mensagem Padrão"
        End If
    End If
    xFaseErro = 18
    Exit Function

FileError:
    Dim xNumeroErro As String
    
    Call CriaLogTEF(Date & " " & Time & " GerenciadorPadraoAtivo: Fase do erro: " & xFaseErro & " - ErroNúmero: " & xNumeroErro & " - ErroTexto: " & Error)
    xNumeroErro = Err
    If xNumeroErro = "76" Then
        MsgBox "Gerenciador Padrão não está instalado neste computador!", vbInformation, "Transação com TEF não aceita!"
        Exit Function
    End If
    MsgBox "ERrO não identificado!" & Chr(10) & Error & Chr(10) & Err, vbInformation, "Rotina: GerenciadorPadraoAtivo"
    Exit Function
End Function
Function SolicitacaoNCN() As Boolean
    Dim i As Integer
    Dim xValor As String
    Dim xNomeRede As String
    Dim xNumeroTransacao As String
    Dim xFinalizacao As String
    Dim xString As String
    Dim xStringNSU As String
    Dim xNomeArquivo As String        'Arquivo a Ser Criado "ATV"
    Dim xNomeArquivoResp As String    'Arquivo de Resposta
    Dim xNomeArquivoResp001 As String 'Arquivo de Resposta
    Dim xNomeArquivoTemp As String    'Arquivo Temporário
    
    Call CriaLogTEF(Date & " " & Time & " Iniciada SolicitacaoNCN ")
    
    'On Error GoTo FileError
    
    SolicitacaoNCN = False
'    If UCase(gBandeira) = "HIPERTEF" Or UCase(gBandeira) = "HIPERCARD" Then
'        xNomeArquivo = "C:\HiperTEF\REQ\IntPos.001"
'        xNomeArquivoResp = "C:\HiperTEF\RESP\IntPos.STS"
'        xNomeArquivoResp001 = "C:\HiperTEF\RESP\IntPos.001"
'        xNomeArquivoTemp = "C:\HiperTEF\REQ\IntPos.tmp"
'    ElseIf UCase(gBandeira) = "TECBAN" Then
'        xNomeArquivo = "C:\TEF_DISC\REQ\IntPos.001"
'        xNomeArquivoResp = "C:\TEF_DISC\RESP\IntPos.STS"
'        xNomeArquivoResp001 = "C:\TEF_DISC\RESP\IntPos.001"
'        xNomeArquivoTemp = "C:\TEF_DISC\REQ\IntPos.tmp"
'    ElseIf UCase(gBandeira) = "TCSMART" Then
'        xNomeArquivo = "C:\TCS\TX\IntTCS.001"
'        xNomeArquivoResp = "C:\TCS\RX\IntTCS.RET"
'        xNomeArquivoResp001 = "C:\TCS\RX\IntTcs.001"
'        xNomeArquivoTemp = "C:\TCS\TX\IntTCS.tmp"
'    ElseIf UCase(gBandeira) = "SMARTEF" Then
'        xNomeArquivo = "C:\SMARTEF\REQ\IntPos.001"
'        xNomeArquivoResp = "C:\SMARTEF\RESP\IntPos.STS"
'        xNomeArquivoResp001 = "C:\SMARTEF\RESP\IntPos.001"
'        xNomeArquivoTemp = "C:\SMARTEF\REQ\IntPos.tmp"
'    ElseIf UCase(gBandeira) = "SUPERTEF" Then
'        xNomeArquivo = "C:\SUPERTEF\REQ\IntPos.001"
'        xNomeArquivoResp = "C:\SUPERTEF\RESP\IntPos.STS"
'        xNomeArquivoResp001 = "C:\SUPERTEF\RESP\IntPos.001"
'        xNomeArquivoTemp = "C:\SUPERTEF\REQ\IntPos.tmp"
'    ElseIf UCase(gBandeira) = "HIPERTEF" Then
'        xNomeArquivo = "C:\HiperTEF\REQ\IntPos.001"
'        xNomeArquivoResp = "C:\HiperTEF\RESP\IntPos.STS"
'        xNomeArquivoResp001 = "C:\HiperTEF\RESP\IntPos.001"
'        xNomeArquivoTemp = "C:\HiperTEF\REQ\IntPos.tmp"
'    ElseIf UCase(gBandeira) = "PAGCARD" Then
'        xNomeArquivo = "C:\CardTech_NEUS\REQ\IntPos.001"
'        xNomeArquivoResp = "C:\CardTech_NEUS\RESP\IntPos.STS"
'        xNomeArquivoResp001 = "C:\CardTech_NEUS\RESP\IntPos.001"
'        xNomeArquivoTemp = "C:\CardTech_NEUS\REQ\IntPos.tmp"
'    Else
'        xNomeArquivo = "C:\TEF_DIAL\REQ\IntPos.001"
'        xNomeArquivoResp = "C:\TEF_DIAL\RESP\IntPos.STS"
'        xNomeArquivoResp001 = "C:\TEF_DIAL\RESP\IntPos.001"
'        xNomeArquivoTemp = "C:\TEF_DIAL\REQ\IntPos.tmp"
'    End If
    
    
    If UCase(gBandeira) = "TCSMART" Then
        xNomeArquivo = gDiretorioReq & "IntTCS.001"
        xNomeArquivoResp = gDiretorioResp & "IntTCS.RET"
        xNomeArquivoResp001 = gDiretorioResp & "IntTcs.001"
        xNomeArquivoTemp = gDiretorioReq & "IntTCS.tmp"
    Else
        xNomeArquivo = gDiretorioReq & "IntPos.001"
        xNomeArquivoResp = gDiretorioResp & "IntPos.STS"
        xNomeArquivoResp001 = gDiretorioResp & "IntPos.001"
        xNomeArquivoTemp = gDiretorioReq & "IntPos.tmp"
    End If
    
    LoopGerenciadorPadrao
    
    'Deletar o backup do "IntPos.001" em c:\Backup_Sgp
    If gArqTxt.FileExists("C:\Backup_SGP\IntPos.001") Then
        Call ExcluirArquivo("C:\Backup_SGP\IntPos.001")
    End If
    
    If CarregaMensagemTEF(False, False, False) Then
        xValor = lCampo003
        xNomeRede = lCampo010
        xNumeroTransacao = lCampo012
        xFinalizacao = lCampo027
        Call ExcluirArquivo(xNomeArquivoResp001)
    Else
        MsgBox "O arquivo IntPos.001 foi deletado de forma não identificada!", vbInformation, "Rotina: SolicitacaoNCN"
        Exit Function
    End If
    If UCase(gBandeira) = "TCSMART" Then
        MsgBox "Pendência Ticket Car Smart Resolvida!", vbInformation, "Pendência"
        SolicitacaoNCN = True
        Exit Function
    End If
    
    xString = ""
    'If gBandeira = "VISANET" Then
    If xValor <> "" Then
        xString = Chr(10) & Chr(10) & "Valor: " & xValor
    End If
    'Else
    '    If Val(xValor) > 0 Then
    '        i = Len(xValor)
    '        If Len(xValor) >= 3 Then
    '            xString = Chr(10) & Chr(10) & "Valor: " & Format(Mid(xValor, 1, i - 2) & "," & Mid(xValor, i - 1, 2), "###,###,##0.00")
    '        Else
    '            xString = Chr(10) & Chr(10) & "Valor: 0," & xValor
    '        End If
    '    End If
    'End If
    
    If lCampo009 = "FF" Then
        Exit Function
    End If
    
    gNumeroControleSolicitacao = lControleSolicitacao.ProximaSolicitacaoTEF(1)
    Set gArquivo = gArqTxt.CreateTextFile(xNomeArquivoTemp)
    gArquivo.WriteLine ("000-000 = NCN")
    gArquivo.WriteLine ("001-000 = " & Format(gNumeroControleSolicitacao, "0000000000"))
    gArquivo.WriteLine ("010-000 = " & xNomeRede)
    gArquivo.WriteLine ("012-000 = " & xNumeroTransacao)
    gArquivo.WriteLine ("027-000 = " & xFinalizacao)
    If gTipoDesconto = "POSTOAKI" Then
        gArquivo.WriteLine ("600-000 = " & gCodigoColaborador)
        gArquivo.WriteLine ("601-000 = " & gNomeColaborador)
        gArquivo.WriteLine ("602-000 = " & gAvaliacaoColaborador)
    End If
    gArquivo.WriteLine ("999-999 = 0")
    gArquivo.Close

    If gArqTxt.FileExists(xNomeArquivoTemp) Then
        Call RenomearArquivo(xNomeArquivoTemp, xNomeArquivo)
    End If
    
    lHoraInicial = Time
    Do Until DateDiff("s", lHoraInicial, Time) >= 7
        If gArqTxt.FileExists(xNomeArquivoResp) Then
            'BUG:XP DÁ UM TEMPO PARA LIBERAR ARQUIVO PARA LEITURA
            lHoraInicial = Time
            Do Until DateDiff("s", lHoraInicial, Time) >= 2
            Loop
            Exit Do
        End If
        DoEvents
    Loop
    If gArqTxt.FileExists(xNomeArquivoResp) Then
        Set gArquivo = gArqTxt.OpenTextFile(xNomeArquivoResp, ForReading)
        lCampo001 = False
        lString = gArquivo.ReadLine
        If lString = "000-000 = NCN" Then
            Do Until gArquivo.AtEndOfStream
                lString = gArquivo.ReadLine
                If Mid(lString, 1, 7) = "001-000" Then
                    If Mid(lString, 11, 10) = gNumeroControleSolicitacao Then
                        lCampo001 = True
                    End If
                End If
            Loop
        End If
        gArquivo.Close
        Call ExcluirArquivo(xNomeArquivoResp)
        If lCampo001 Then
        
            SolicitacaoNCN = True
            xStringNSU = ""
            If Len(xNumeroTransacao) > 0 Then
                xStringNSU = Chr(10) & Chr(10) & "NSU: " & xNumeroTransacao
            End If
            Call CriaLogTEF(Date & " " & Time & " Transação TEF Cancelada: NCN " & rtxt_mensagem.Text & " - Bandeira=" & xNomeRede & " - Valor=" & xValor & " - Numero da Transação=" & xNumeroTransacao)
            MsgBox "Última Transação TEF foi Cancelada" & Chr(10) & Chr(10) & "Rede: " & xNomeRede & xStringNSU & xString, vbInformation, "Última transação TEF foi cancelada."
            'Nao Tentar fechar um possível Relatório Gerencial Aqui.
            'Testar se está aberto e fecha-lo no cupom fiscal
            'If gImpBematech Then
            '    BemaRetorno = Bematech_FI_FechaRelatorioGerencial
            'End If
        Else
            MsgBox "Numero de controle de solicitação está diferente", vbInformation, "SolicitacaoNCN"
        End If
    Else
        'Como não teve respostas em 7 segundos
        'Deleta Arquivo de Requisicao
        If gArqTxt.FileExists(xNomeArquivo) Then
            Call ExcluirArquivo(xNomeArquivo)
        End If
        MsgBox "TEF Não Está Ativo!", vbInformation, "Mensagem Padrão"
    End If
End Function
Private Function LoopNumeroControleSolicitacao(ByVal pNomeArquivo As String) As Boolean
    Dim xSaiLoop As Boolean
    Dim xCampo001Igual As Boolean
    
    xSaiLoop = False
    xCampo001Igual = False
    LoopNumeroControleSolicitacao = False
    Do Until xSaiLoop = True
        Do Until gArqTxt.FileExists(pNomeArquivo)
            DoEvents
        Loop
        'BUG:XP DÁ UM TEMPO PARA LIBERAR ARQUIVO PARA LEITURA
        lHoraInicial = Time
        Do Until DateDiff("s", lHoraInicial, Time) >= 2
        Loop
        Set gArquivo = gArqTxt.OpenTextFile(pNomeArquivo, ForReading)
        Do Until gArquivo.AtEndOfStream
            lString = gArquivo.ReadLine
            If UCase(gBandeira) = "TCSMART" Then
                LoopNumeroControleSolicitacao = True
                xSaiLoop = True
                Exit Do
            Else
                If Mid(lString, 1, 7) = "001-000" Then
                    If Mid(lString, 11, 10) = gNumeroControleSolicitacao Then
                        LoopNumeroControleSolicitacao = True
                        xCampo001Igual = True
                        xSaiLoop = True
                    End If
                    Exit Do
                End If
            End If
        Loop
        gArquivo.Close
        If xCampo001Igual = False Then
            ExcluirArquivo (pNomeArquivo)
        End If
    Loop
End Function
Private Function LoopGerenciadorPadrao() As Boolean
    LoopGerenciadorPadrao = False
    
    Do Until GerenciadorPadraoAtivo(True, True)
        LoopGerenciadorPadrao = True
    Loop
    LoopGerenciadorPadrao = True
End Function
Function SolicitacaoADM() As Boolean
    Dim i As Integer
    Dim xNomeArquivo As String        'Arquivo a Ser Criado "ATV"
    Dim xNomeArquivoResp As String    'Arquivo de Resposta
    Dim xNomeArquivoResp001 As String 'Arquivo de Resposta
    Dim xNomeArquivoTemp As String    'Arquivo Temporário
    Dim retval As Long
    Dim xFaseErro As Integer
    
    On Error GoTo FileError
    
    SolicitacaoADM = False
    lCampo001 = False
    xFaseErro = 1
'    If UCase(gBandeira) = "TECBAN" Then
'        xNomeArquivo = "C:\TEF_DISC\REQ\IntPos.001"
'        xNomeArquivoResp = "C:\TEF_DISC\RESP\IntPos.STS"
'        xNomeArquivoResp001 = "C:\TEF_DISC\RESP\IntPos.001"
'        xNomeArquivoTemp = "C:\TEF_DISC\REQ\IntPos.tmp"
'    ElseIf UCase(gBandeira) = "TCSMART" Then
'        xNomeArquivo = "C:\TCS\TX\IntTcs.001"
'        xNomeArquivoResp = "C:\TCS\RX\IntTcs.RET"
'        xNomeArquivoResp001 = "C:\TCS\RX\IntTcs.001"
'        xNomeArquivoTemp = "C:\TCS\TX\IntTcs.tmp"
'    ElseIf UCase(gBandeira) = "SMARTEF" Then
'        xNomeArquivo = "C:\SMARTEF\REQ\IntPos.001"
'        xNomeArquivoResp = "C:\SMARTEF\RESP\IntPos.STS"
'        xNomeArquivoResp001 = "C:\SMARTEF\RESP\IntPos.001"
'        xNomeArquivoTemp = "C:\SMARTEF\REQ\IntPos.tmp"
'    ElseIf UCase(gBandeira) = "SUPERTEF" Then
'        xNomeArquivo = "C:\SUPERTEF\REQ\IntPos.001"
'        xNomeArquivoResp = "C:\SUPERTEF\RESP\IntPos.STS"
'        xNomeArquivoResp001 = "C:\SUPERTEF\RESP\IntPos.001"
'        xNomeArquivoTemp = "C:\SUPERTEF\REQ\IntPos.tmp"
'    ElseIf UCase(gBandeira) = "HIPERTEF" Then
'        xNomeArquivo = "C:\HiperTEF\REQ\IntPos.001"
'        xNomeArquivoResp = "C:\HiperTEF\RESP\IntPos.STS"
'        xNomeArquivoResp001 = "C:\HiperTEF\RESP\IntPos.001"
'        xNomeArquivoTemp = "C:\HiperTEF\REQ\IntPos.tmp"
'    ElseIf UCase(gBandeira) = "PAGCARD" Then
'        xNomeArquivo = "C:\CardTech_NEUS\REQ\IntPos.001"
'        xNomeArquivoResp = "C:\CardTech_NEUS\RESP\IntPos.STS"
'        xNomeArquivoResp001 = "C:\CardTech_NEUS\RESP\IntPos.001"
'        xNomeArquivoTemp = "C:\CardTech_NEUS\REQ\IntPos.tmp"
'    Else
'        xNomeArquivo = "C:\TEF_DIAL\REQ\IntPos.001"
'        xNomeArquivoResp = "C:\TEF_DIAL\RESP\IntPos.STS"
'        xNomeArquivoResp001 = "C:\TEF_DIAL\RESP\IntPos.001"
'        xNomeArquivoTemp = "C:\TEF_DIAL\REQ\IntPos.tmp"
'    End If
    If UCase(gBandeira) = "TCSMART" Then
        xNomeArquivo = gDiretorioReq & "IntTcs.001"
        xNomeArquivoResp = gDiretorioResp & "IntTcs.RET"
        xNomeArquivoResp001 = gDiretorioResp & "IntTcs.001"
        xNomeArquivoTemp = gDiretorioReq & "IntTcs.tmp"
    Else
        xNomeArquivo = gDiretorioReq & "IntPos.001"
        xNomeArquivoResp = gDiretorioResp & "IntPos.STS"
        xNomeArquivoResp001 = gDiretorioResp & "IntPos.001"
        xNomeArquivoTemp = gDiretorioReq & "IntPos.tmp"
    End If
    
    'ADM
    'Se existir arquivos, deleta.
    xFaseErro = 4
    If gArqTxt.FileExists(xNomeArquivoTemp) Then
        Call ExcluirArquivo(xNomeArquivoTemp)
    End If
    xFaseErro = 6
    If gArqTxt.FileExists(xNomeArquivo) Then
        Call ExcluirArquivo(xNomeArquivo)
    End If
    
    
    'Gera Arquivo
    xFaseErro = 8
    Set gArquivo = gArqTxt.CreateTextFile(xNomeArquivoTemp)
    xFaseErro = 10
    If UCase(gBandeira) = "TCSMART" Then
        gArquivo.WriteLine ("500-000 = ADM")
        gArquivo.WriteLine ("501-000 = " & Format(gNumeroControleSolicitacao, "0000000000"))
    Else
        gArquivo.WriteLine ("000-000 = ADM")
        gArquivo.WriteLine ("001-000 = " & Format(gNumeroControleSolicitacao, "0000000000"))
    End If
    If gTipoDesconto = "POSTOAKI" Then
        gArquivo.WriteLine ("600-000 = " & gCodigoColaborador)
        gArquivo.WriteLine ("601-000 = " & gNomeColaborador)
        gArquivo.WriteLine ("602-000 = " & gAvaliacaoColaborador)
    End If
    gArquivo.WriteLine ("999-999 = 0")
    gArquivo.Close
    
    xFaseErro = 20
    If gArqTxt.FileExists(xNomeArquivoTemp) Then
        Call RenomearArquivo(xNomeArquivoTemp, xNomeArquivo)
        Call CriaLogTefEspecial("SolicitacaoADM - 001 Foi criado o arquivo xNomeArquivo: " & xNomeArquivo & " - gNumeroControleSolicitacao: " & gNumeroControleSolicitacao & " - gTipoDesconto: ->" & gTipoDesconto & "<- gBandeira: " & gBandeira)
    End If
    
    xFaseErro = 22
    If UCase(gBandeira) = "TCSMART" Then
        retval = Shell("C:\TCS\tcs.exe", vbNormalFocus)
        Do Until gArqTxt.FileExists(xNomeArquivoResp)
            DoEvents
        Loop
    End If
    'Aguarda 7 segundos
    xFaseErro = 26
    lHoraInicial = Time
    Do Until DateDiff("s", lHoraInicial, Time) >= 7
        If gArqTxt.FileExists(xNomeArquivoResp) Then
            If UCase(gBandeira) <> "TEFCERRADO" Then
                'BUG:XP DÁ UM TEMPO PARA LIBERAR ARQUIVO PARA LEITURA
                lHoraInicial = Time
                Do Until DateDiff("s", lHoraInicial, Time) >= 2
                Loop
            End If
            Exit Do
        End If
        DoEvents
    Loop
    
    xFaseErro = 40
    Call CriaLogTefEspecial("SolicitacaoADM - 002 Verifica se existe arquivo de Retorno. xNomeArquivoResp: " & xNomeArquivoResp)
    If Not gArqTxt.FileExists(xNomeArquivoResp) Then
        MsgBox "Gerenciador Padrão não está ativo, e será ativado automaticamente!", vbInformation, "Mensagem Padrão"
        Call AtivaGerenciadorPadrao
        'Aguarda 7 segundos
        lHoraInicial = Time
        Do Until DateDiff("s", lHoraInicial, Time) >= 7
            If gArqTxt.FileExists(xNomeArquivoResp) Then
                If UCase(gBandeira) <> "TEFCERRADO" Then
                    'BUG:XP DÁ UM TEMPO PARA LIBERAR ARQUIVO PARA LEITURA
                    lHoraInicial = Time
                    Do Until DateDiff("s", lHoraInicial, Time) >= 2
                    Loop
                End If
                Exit Do
            End If
            DoEvents
        Loop
    End If
    
    xFaseErro = 50
    If gArqTxt.FileExists(xNomeArquivoResp) Then
        Call CriaLogTefEspecial("SolicitacaoADM - 003 Foi localizado o arquivo de Retorno. xNomeArquivoResp: " & xNomeArquivoResp)
        xFaseErro = 51
        If UCase(gBandeira) = "TEFCERRADO" Then
            'BUG:XP DÁ UM TEMPO PARA LIBERAR ARQUIVO PARA LEITURA
            lHoraInicial = Time
            Do Until DateDiff("s", lHoraInicial, Time) >= 2
            Loop
        End If
        xFaseErro = 52
        Set gArquivo = gArqTxt.OpenTextFile(xNomeArquivoResp, ForReading)
        xFaseErro = 53
        lString = gArquivo.ReadLine
        Call CriaLogTefEspecial("SolicitacaoADM - 004 Primeira linha de retorno. lString: " & lString)
        xFaseErro = 54
        If UCase(gBandeira) = "TCSMART" Then
            xFaseErro = 60
            If lString = "517-000 = 000" Then
                gArquivo.Close
                Set gArquivo = gArqTxt.OpenTextFile(xNomeArquivoResp001, ForReading)
                Do Until gArquivo.AtEndOfStream
                    lString = gArquivo.ReadLine
                    Call CriaLogTefEspecial("SolicitacaoADM - 005 Sequencia TCS 517-000 retorno. lString: " & lString)
                    'Verifica se o número do controle da solicitação é igual
                    If Mid(lString, 1, 7) = "501-000" Then
                        If Mid(lString, 11, 10) = gNumeroControleSolicitacao Then
                            lCampo001 = True
                            Exit Do
                        End If
                    End If
                Loop
            Else
                lString = gArquivo.ReadLine
                Call CriaLogTefEspecial("SolicitacaoADM - 005 Sequencia TCS linha   retorno. lString: " & lString)
                lString = Mid(lString, 12, Len(lString) - 12)
            End If
        Else
            xFaseErro = 70
            If lString = "000-000 = ADM" Then
                Do Until gArquivo.AtEndOfStream
                    lString = gArquivo.ReadLine
                    Call CriaLogTefEspecial("SolicitacaoADM - 005 Sequencia linha   retorno. lString: " & lString)
                    'Verifica se o número do controle da solicitação é igual
                    If Mid(lString, 1, 7) = "001-000" Then
                        If Mid(lString, 11, 10) = gNumeroControleSolicitacao Then
                            Call CriaLogTefEspecial("SolicitacaoADM - 006 Campo lCampo001 passa a ser verdadeiro. gNumeroControleSolicitacao: " & gNumeroControleSolicitacao)
                            lCampo001 = True
                            Exit Do
                        End If
                    End If
                Loop
            End If
        End If
        xFaseErro = 80
        gArquivo.Close
        
        xFaseErro = 81
        Call ExcluirArquivo(xNomeArquivoResp)
        xFaseErro = 82
        If lCampo001 Then
            SolicitacaoADM = True
        Else
            xFaseErro = 85
            If UCase(gBandeira) = "TCSMART" Then
                MsgBox lString, vbInformation, "Solicitação ADM"
            Else
                MsgBox "Numero de controle de solicitação está diferente", vbInformation, "SolicitacaoADM"
            End If
        End If
    Else
        'Como não teve respostas em 7 segundos
        'Deleta Arquivo de Requisicao
        Call CriaLogTefEspecial("SolicitacaoADM - 007 Sem retorno em 7 segundos.")
        xFaseErro = 90
        If gArqTxt.FileExists(xNomeArquivo) Then
            xFaseErro = 91
            Call ExcluirArquivo(xNomeArquivo)
        End If
        xFaseErro = 92
        MsgBox "Gerenciador Padrão não está ativo!", vbInformation, "Mensagem Padrão"
    End If
    Exit Function
FileError:
    Call CriaLogTefEspecial("SolicitacaoADM - 008 Erro Número: " & Err & " - ErroTexto: " & Error)
    Call CriaLogTEF(Date & " " & Time & " SolicitacaoADM: " & " - ErroNúmero: " & Err & " - ErroTexto: " & Error)
    If Err = 76 Then
        MsgBox "Gerenciador Padrão não está instalado neste computador!" & vbCrLf & "xFaseErro:" & xFaseErro, vbInformation, "Transação com TEF não aceita!"
        Exit Function
    End If
    MsgBox "ERRO não identificado!" & Chr(10) & Error & vbCrLf & "xFaseErro:" & xFaseErro, vbInformation, "Rotina: SolicitacaoADM"
    MsgBox "Arquivo: " & xNomeArquivoResp
    Exit Function
End Function
Function SolicitacaoDeCompra() As Boolean
    Dim i As Integer
    Dim xNomeArquivo As String     'Arquivo a Ser Criado "CRT"
    Dim xNomeArquivoResp As String 'Arquivo de Resposta
    Dim xNomeArquivoResp001 As String
    Dim xNomeArquivoTemp As String 'Arquivo Temporário
    Dim xDadosTCS As Variant
    Dim retval As Long
    
    On Error GoTo FileError
    
    SolicitacaoDeCompra = False
    lCampo001 = False
    If UCase(gBandeira) = "TCSMART" Then
        xNomeArquivo = gDiretorioReq & "IntTcs.001"
        xNomeArquivoResp = gDiretorioResp & "IntTcs.RET"
        xNomeArquivoResp001 = gDiretorioResp & "IntTcs.001"
        xNomeArquivoTemp = gDiretorioReq & "IntTcs.tmp"
    Else
        xNomeArquivo = gDiretorioReq & "IntPos.001"
        xNomeArquivoResp = gDiretorioResp & "IntPos.STS"
        xNomeArquivoResp001 = gDiretorioResp & "IntPos.001"
        xNomeArquivoTemp = gDiretorioReq & "IntPos.tmp"
    End If
    
    
    'CRT
    'Se existir arquivos, deleta.
    Call ExcluirArquivo(xNomeArquivoTemp)
    Call ExcluirArquivo(xNomeArquivo)
    
    If UCase(gBandeira) = "TCSMART" Then
    Else
        gNumeroControleSolicitacao = lControleSolicitacao.ProximaSolicitacaoTEF(1)
    End If
    'Gera Arquivo
    Call CriarArquivo(xNomeArquivoTemp)
    'Set gArquivo = gArqTxt.CreateTextFile(xNomeArquivoTemp)
    If UCase(gBandeira) = "TCSMART" Then
        gArquivo.WriteLine ("500-000 = CRT")
        gArquivo.WriteLine ("501-000 = " & Format(gNumeroControleSolicitacao, "0000000000"))
        gArquivo.WriteLine ("502-000 = 0001")
        gArquivo.WriteLine ("503-000 = " & gContadorNaoFiscal)
        gArquivo.WriteLine ("504-000 = " & Format(gNumeroCupom + 1, "000000"))
        gArquivo.WriteLine ("505-000 = 0001")
        For i = 0 To UBound(gDadosTCS) - 1
            xDadosTCS = Split(gDadosTCS(i), "|@|")
            gArquivo.WriteLine ("506-000 = " & xDadosTCS(2))
            gArquivo.WriteLine ("507-000 = " & xDadosTCS(1))
            gArquivo.WriteLine ("509-000 = " & xDadosTCS(0))
            gArquivo.WriteLine ("509-001 = " & xDadosTCS(3))
            gArquivo.WriteLine ("510-000 = " & xDadosTCS(4))
            gArquivo.WriteLine ("518-000 = " & xDadosTCS(5))
            gArquivo.WriteLine ("519-000 = " & xDadosTCS(6))
        Next
        If gLegislacaoPermiteIssEcf = True Then
            gArquivo.WriteLine ("520-000 = 1")
            gArquivo.WriteLine ("521-000 = 0000")
            gArquivo.WriteLine ("522-000 = 00")
        Else
            gArquivo.WriteLine ("520-000 = 0")
            gArquivo.WriteLine ("521-000 = " & Mid(gContadorNaoFiscal, 3, 4))
            gArquivo.WriteLine ("522-000 = " & Format(gCodigoTcsEcf, "00"))
        End If
        gArquivo.WriteLine ("999-999 = 0")
    Else
        If gConsultaCheque Then
            gArquivo.WriteLine ("000-000 = CHQ")
        Else
            gArquivo.WriteLine ("000-000 = CRT")
        End If
        gArquivo.WriteLine ("001-000 = " & Format(gNumeroControleSolicitacao, "0000000000"))
        gArquivo.WriteLine ("002-000 = " & Format(gNumeroCupom + 1, "000000"))
        lString = Format(fValidaValor(gValorRecebido), "#########0.00")
        i = Len(lString)
        lString = Mid(lString, 1, i - 3) & Mid(lString, i - 1, 2)
        gArquivo.WriteLine ("003-000 = " & lString)
        'teste dominio automacao 777
        'gArquivo.WriteLine ("777-777 = teste redecard")
        If UCase(gBandeira) = "TEFCERRADO" Then
            For i = 0 To UBound(gDadosProdutos) - 1
                If Len(RetiraString(2, gDadosProdutos(i))) > 0 Then
                    gArquivo.WriteLine ("550-" & Format(i, "000") & " = " & Format(Val(RetiraString(1, gDadosProdutos(i))), "000000"))
                    gArquivo.WriteLine ("551-" & Format(i, "000") & " = " & RetiraString(2, gDadosProdutos(i)))
                    gArquivo.WriteLine ("552-" & Format(i, "000") & " = " & Format(CCur(RetiraString(3, gDadosProdutos(i))) * 100, "000000000000"))
                    gArquivo.WriteLine ("553-" & Format(i, "000") & " = " & Format(CCur(RetiraString(4, gDadosProdutos(i))) * 100, "000000000000"))
                    gArquivo.WriteLine ("554-" & Format(i, "000") & " = " & RetiraString(5, gDadosProdutos(i)))
                End If
            Next
        End If
        If gTipoDesconto = "POSTOAKI" Then
            gArquivo.WriteLine ("600-000 = " & gCodigoColaborador)
            gArquivo.WriteLine ("601-000 = " & gNomeColaborador)
            gArquivo.WriteLine ("602-000 = " & gAvaliacaoColaborador)
        End If
        gArquivo.WriteLine ("999-999 = 0")
    End If
    gArquivo.Close
    
    Call RenomearArquivo(xNomeArquivoTemp, xNomeArquivo)
    If UCase(gBandeira) = "TCSMART" Then
        retval = Shell("C:\TCS\tcs.exe", vbNormalFocus)
        Do Until gArqTxt.FileExists(xNomeArquivoResp)
            DoEvents
        Loop
    End If
    'Aguarda 7 segundos
    lHoraInicial = Time
    Do Until DateDiff("s", lHoraInicial, Time) >= 7
        If gArqTxt.FileExists(xNomeArquivoResp) Then
            If UCase(gBandeira) <> "TEFCERRADO" Then
                'BUG:XP DÁ UM TEMPO PARA LIBERAR ARQUIVO PARA LEITURA
                Call AguardaTempo(2)
            End If
            Exit Do
        End If
        DoEvents
    Loop
    
    
    
    If gArqTxt.FileExists(xNomeArquivoResp) Then
        Call AbrirArquivo(xNomeArquivoResp, "LEITURA")
        lString = gArquivo.ReadLine
        If UCase(gBandeira) = "TCSMART" Then
            If lString = "517-000 = 000" Then
                gArquivo.Close
                Call AbrirArquivo(xNomeArquivoResp001, "LEITURA")
                Do Until gArquivo.AtEndOfStream
                    lString = gArquivo.ReadLine
                    If Mid(lString, 1, 7) = "501-000" Then
                        If Mid(lString, 11, 10) = gNumeroControleSolicitacao Then
                            lCampo001 = True
                            Exit Do
                        End If
                    End If
                Loop
            Else
                lString = gArquivo.ReadLine
                lString = Mid(lString, 12, Len(lString) - 12)
            End If
        Else
            If lString = "000-000 = CRT" Or lString = "000-000 = CHQ" Then
                Do Until gArquivo.AtEndOfStream
                    lString = gArquivo.ReadLine
                    If Mid(lString, 1, 7) = "001-000" Then
                        If Mid(lString, 11, 10) = gNumeroControleSolicitacao Then
                            lCampo001 = True
                        End If
                    End If
                Loop
            End If
        End If
        gArquivo.Close
        
        Call ExcluirArquivo(xNomeArquivoResp)
        If lCampo001 Then
            SolicitacaoDeCompra = True
            Call CriaLogTEF(Date & " " & Time & " SolicitacaoDeCompra: " & rtxt_mensagem.Text & " - Bandeira=" & gBandeira & " - Valor=" & gValorRecebido & " - Numero do Controle de Solicitacao=" & gNumeroControleSolicitacao)
        Else
            If UCase(gBandeira) = "TCSMART" Then
                MsgBox lString, vbInformation, "SolicitacaoDeCompra"
            Else
                MsgBox "Numero de controle de solicitação está diferente", vbInformation, "SolicitacaoDeCompra"
            End If
        End If
    Else
        'Como não teve respostas em 7 segundos
        'Deleta Arquivo de Requisicao
        Call ExcluirArquivo(xNomeArquivo)
        MsgBox "TEF Não Está Ativo!", vbInformation, "Mensagem Padrão"
    End If
    Exit Function
FileError:
    Call CriaLogTEF("Erro SolicitacaoDeCompra: " & " - ErroNúmero: " & Err & " - ErroTexto: " & Error & " - às: " & Time)
    MsgBox "ERRO não identificado!", vbInformation, "Rotina: SolicitacaoDeCompra"
    Exit Function
End Function
Function SolicitacaoDeDesconto() As Boolean
    Dim i As Integer
    Dim xNomeArquivo As String     'Arquivo a Ser Criado "DSC"
    Dim xNomeArquivoResp As String 'Arquivo de Resposta
    Dim xNomeArquivoResp001 As String
    Dim xNomeArquivoTemp As String 'Arquivo Temporário
    Dim xDadosTCS As Variant
    Dim retval As Long
    
    On Error GoTo FileError
    
    SolicitacaoDeDesconto = False
    lCampo001 = False
    xNomeArquivo = gDiretorioReq & "IntPos.001"
    xNomeArquivoResp = gDiretorioResp & "IntPos.STS"
    xNomeArquivoResp001 = gDiretorioResp & "IntPos.001"
    xNomeArquivoTemp = gDiretorioReq & "IntPos.tmp"
    
    
    'DSC
    'Se existir arquivos, deleta.
    Call ExcluirArquivo(xNomeArquivoTemp)
    Call ExcluirArquivo(xNomeArquivo)
    
    gNumeroControleSolicitacao = lControleSolicitacao.ProximaSolicitacaoTEF(1)
    'Gera Arquivo
    Call CriarArquivo(xNomeArquivoTemp)
    'Set gArquivo = gArqTxt.CreateTextFile(xNomeArquivoTemp)
    gArquivo.WriteLine ("000-000 = DSC")
    gArquivo.WriteLine ("001-000 = " & Format(gNumeroControleSolicitacao, "0000000000"))
    gArquivo.WriteLine ("002-000 = " & Format(gNumeroCupom + 1, "000000"))
    lString = Format(fValidaValor(gValorRecebido), "#########0.00")
    i = Len(lString)
    lString = Mid(lString, 1, i - 3) & Mid(lString, i - 1, 2)
    gArquivo.WriteLine ("003-000 = " & lString)
    For i = 0 To UBound(gDadosProdutos) - 1
        If Len(RetiraString(2, gDadosProdutos(i))) > 0 Then
            gArquivo.WriteLine ("550-" & Format(i, "000") & " = " & Format(Val(RetiraString(1, gDadosProdutos(i))), "000000"))
            gArquivo.WriteLine ("551-" & Format(i, "000") & " = " & RetiraString(2, gDadosProdutos(i)))
            gArquivo.WriteLine ("552-" & Format(i, "000") & " = " & Format(CCur(RetiraString(3, gDadosProdutos(i))) * 100, "000000000000"))
            gArquivo.WriteLine ("553-" & Format(i, "000") & " = " & Format(CCur(RetiraString(4, gDadosProdutos(i))) * 100, "000000000000"))
            gArquivo.WriteLine ("554-" & Format(i, "000") & " = " & RetiraString(5, gDadosProdutos(i)))
        End If
    Next
    If gTipoDesconto = "POSTOAKI" Then
        gArquivo.WriteLine ("600-000 = " & gCodigoColaborador)
        gArquivo.WriteLine ("601-000 = " & gNomeColaborador)
        gArquivo.WriteLine ("602-000 = " & gAvaliacaoColaborador)
    End If
    gArquivo.WriteLine ("950-000 = " & gTipoDesconto)
    gArquivo.WriteLine ("951-000 = " & gNumeroAutorizacaoPostoAki)
    If gTrocaOleo = True Then
        gArquivo.WriteLine ("952-000 = 1")
    Else
        gArquivo.WriteLine ("952-000 = 0")
    End If
    If gPontuacao = True Then
        gArquivo.WriteLine ("953-000 = 1")
    Else
        gArquivo.WriteLine ("953-000 = 0")
    End If
    gArquivo.WriteLine ("999-999 = 0")
    gArquivo.Close
    
    Call RenomearArquivo(xNomeArquivoTemp, xNomeArquivo)
    'Aguarda 7 segundos
    lHoraInicial = Time
    Do Until DateDiff("s", lHoraInicial, Time) >= 7
        If gArqTxt.FileExists(xNomeArquivoResp) Then
            'BUG:XP DÁ UM TEMPO PARA LIBERAR ARQUIVO PARA LEITURA
            Call AguardaTempo(2)
            Exit Do
        End If
        DoEvents
    Loop
    
    If gArqTxt.FileExists(xNomeArquivoResp) Then
        Call AbrirArquivo(xNomeArquivoResp, "LEITURA")
        lString = gArquivo.ReadLine
        If lString = "000-000 = DSC" Then
            Do Until gArquivo.AtEndOfStream
                lString = gArquivo.ReadLine
                If Mid(lString, 1, 7) = "001-000" Then
                    If Mid(lString, 11, 10) = gNumeroControleSolicitacao Then
                        lCampo001 = True
                    End If
                End If
            Loop
        End If
        gArquivo.Close
        
        Call ExcluirArquivo(xNomeArquivoResp)
        If lCampo001 Then
            SolicitacaoDeDesconto = True
            Call CriaLogTEF(Date & " " & Time & " SolicitacaoDeDesconto: " & rtxt_mensagem.Text & " - Bandeira=" & gBandeira & " - Valor=" & gValorRecebido & " - Numero do Controle de Solicitacao=" & gNumeroControleSolicitacao)
        Else
            MsgBox "Numero de controle de solicitação está diferente", vbInformation, "SolicitacaoDeDesconto"
        End If
    Else
        'Como não teve respostas em 7 segundos
        'Deleta Arquivo de Requisicao
        Call ExcluirArquivo(xNomeArquivo)
        MsgBox "TEF Não Está Ativo!", vbInformation, "Mensagem Padrão"
    End If
    Exit Function
FileError:
    Call CriaLogTEF("Erro SolicitacaoDeDesconto: " & " - ErroNúmero: " & Err & " - ErroTexto: " & Error & " - às: " & Time)
    MsgBox "ERRO não identificado!", vbInformation, "Rotina: SolicitacaoDeDesconto"
    Exit Function
End Function
Function TestaImprimeTEF(ByVal xTipoTEF As String) As Boolean
    Dim xSaiDoLoop As Boolean
    
    On Error GoTo FileError
    
    TestaImprimeTEF = False
    xSaiDoLoop = False
    
    'Bloqueia Teclado
    If gTipoDocumentoFiscal <> "NFCe" Then
        DoEvents
        Call Bematech_FI_IniciaModoTEF
        DoEvents
    End If
    
    
    'Carrega Mensagem
    If Not CarregaMensagemTEF(False, True, False) Then
        Exit Function
    End If
    Call CriaLogTEF(Date & " " & Time & " TestaImprimeTEF: " & rtxt_mensagem.Text & " - TipoTef=" & xTipoTEF)
    
    'Imprime e Testa TEF
    Do Until xSaiDoLoop = True
        If ImprimeTEF(xTipoTEF) Then
            TestaImprimeTEF = True
            xSaiDoLoop = True
        Else
            rtxt_mensagem.Text = "Impressora Não Responde."
            Call CriaLogTEF(Date & " " & Time & " TestaImprimeTEF: " & rtxt_mensagem.Text & " - Impressora nao Responde - TipoTef=" & xTipoTEF)
            rtxt_mensagem.Visible = True
            DoEvents
            'Desbloqueia Teclado
            If gTipoDocumentoFiscal <> "NFCe" Then
                Call Bematech_FI_FinalizaModoTEF
            End If
            If (MsgBox("Impressora Não Responde, Tentar Imprimir Novamente ?  Sim ou Não.", vbQuestion + vbDefaultButton1 + vbYesNo, "Impressora Não Responde!") = vbNo) Then
                Call CriaLogTEF(Date & " " & Time & " TestaImprimeTEF: " & rtxt_mensagem.Text & " - Informou Não Imprimir Novamente - TipoTef=" & xTipoTEF)
                
                'Bloqueia Teclado
                If gTipoDocumentoFiscal <> "NFCe" Then
                    DoEvents
                    Call Bematech_FI_IniciaModoTEF
                    DoEvents
                End If
                
                If gImpBematech Then
                    BemaRetorno = Bematech_FI_FechaComprovanteNaoFiscalVinculado
                    BemaRetorno = Bematech_FI_FechaRelatorioGerencial
                ElseIf gImpQuick Then
                    If EcfQuickEncerraDocumento("", "Cerrado Informatica (62) 3277-1017") Then
                        BemaRetorno = 1
                    Else
                        BemaRetorno = -1
                    End If
                ElseIf gImpElgin Then
                    BemaRetorno = Elgin_FechaComprovanteNaoFiscalVinculado
                    BemaRetorno = Elgin_FechaRelatorioGerencial
                ElseIf gImpDaruma Then
                    BemaRetorno = Daruma_TEF_FechaRelatorio()
                End If
                xSaiDoLoop = True
            Else
                'Bloqueia Teclado
                If gTipoDocumentoFiscal <> "NFCe" Then
                    DoEvents
                    Call Bematech_FI_IniciaModoTEF
                    DoEvents
                End If
                
                rtxt_mensagem.Text = "Aguarde!" & Chr(10) & Chr(10) & "Imprimindo TEF"
                Call CriaLogTEF(Date & " " & Time & " TestaImprimeTEF: " & rtxt_mensagem.Text & " - Informou Sim, Imprimir Novamente - TipoTef=" & xTipoTEF)
                rtxt_mensagem.Visible = True
                
                If gTipoDocumentoFiscal <> "NFCe" Then
                    DoEvents
                    Call Bematech_FI_IniciaModoTEF
                    DoEvents
                End If
                
                If gImpBematech Then
                    BemaRetorno = Bematech_FI_FechaRelatorioGerencial
                ElseIf gImpQuick Then
                    If EcfQuickEncerraDocumento("", "Cerrado Informatica (62) 3277-1017") Then
                        BemaRetorno = 1
                    Else
                        BemaRetorno = -1
                    End If
                ElseIf gImpElgin Then
                    BemaRetorno = Elgin_FechaComprovanteNaoFiscalVinculado
                    'BemaRetorno = Elgin_FechaRelatorioGerencial
                ElseIf gImpDaruma Then
                    BemaRetorno = Daruma_TEF_FechaRelatorio()
                End If
            End If
        End If
        xTipoTEF = "LeituraX"
        If xSaiDoLoop = True Then
            Exit Do
        End If
    Loop
    'Desbloqueia Teclado
    If gTipoDocumentoFiscal <> "NFCe" Then
        Call Bematech_FI_FinalizaModoTEF
    End If
    Exit Function
FileError:
    'Desbloqueia Teclado
    If gTipoDocumentoFiscal <> "NFCe" Then
        Call Bematech_FI_FinalizaModoTEF
    End If
    Call CriaLogTEF(Date & " " & Time & " TestaImprimeTEF: " & " - ErroNúmero: " & Err & " - ErroTexto: " & Error)
End Function
Function TestaImprimeTefADM(ByVal xTipoTEF As String) As Boolean
    Dim xSaiDoLoop As Boolean
    
    On Error GoTo FileError
    
    TestaImprimeTefADM = False
    xSaiDoLoop = False
    
    'Bloqueia Teclado
    If gTipoDocumentoFiscal <> "NFCe" Then
        DoEvents
        Call Bematech_FI_IniciaModoTEF
        DoEvents
    End If
    
    'Carrega Mensagem
    If Not CarregaMensagemTEF(False, True, False) Then
        Exit Function
    End If
    'Imprime e Testa TEF
    Do Until xSaiDoLoop = True
        If ImprimeTefADM(xTipoTEF) Then
            TestaImprimeTefADM = True
            xSaiDoLoop = True
        Else
            rtxt_mensagem.Text = "Impressora Não Responde."
            'Call CriaLogTEF("CarregaMensagemTEF: Mensagem 8: " & rtxt_mensagem.Text & " às: " & Time)
            rtxt_mensagem.Visible = True
            DoEvents
            
            'Desbloqueia Teclado
            If gTipoDocumentoFiscal <> "NFCe" Then
                Call Bematech_FI_FinalizaModoTEF
            End If
            
            If (MsgBox("Impressora Não Responde, Tentar Imprimir Novamente ?  Sim ou Não.", vbQuestion + vbDefaultButton1 + vbYesNo, "Impressora Não Responde!") = vbNo) Then
                
                'Bloqueia Teclado
                If gTipoDocumentoFiscal <> "NFCe" Then
                    DoEvents
                    Call Bematech_FI_IniciaModoTEF
                    DoEvents
                End If
                
                If gImpBematech Then
                    BemaRetorno = Bematech_FI_FechaRelatorioGerencial
                ElseIf gImpQuick Then
                    If EcfQuickEncerraDocumento("", "Cerrado Informatica") Then
                        BemaRetorno = 1
                    Else
                        BemaRetorno = -1
                    End If
                ElseIf gImpElgin Then
                    BemaRetorno = Elgin_FechaRelatorioGerencial
                End If
                
                xSaiDoLoop = True
            Else
                
                'Bloqueia Teclado
                If gTipoDocumentoFiscal <> "NFCe" Then
                    DoEvents
                    Call Bematech_FI_IniciaModoTEF
                    DoEvents
                End If
                
                rtxt_mensagem.Text = "Aguarde!" & Chr(10) & Chr(10) & "Imprimindo TEF" & Chr(10) & "Administrativo."
                'Call CriaLogTEF("CarregaMensagemTEF: Mensagem 9: " & rtxt_mensagem.Text & " às: " & Time)
                rtxt_mensagem.Visible = True
                DoEvents
                If gImpBematech Then
                    BemaRetorno = Bematech_FI_FechaRelatorioGerencial
                ElseIf gImpQuick Then
                    If EcfQuickEncerraDocumento("", "Cerrado Informatica") Then
                        BemaRetorno = 1
                    Else
                        BemaRetorno = -1
                    End If
                ElseIf gImpElgin Then
                    BemaRetorno = Elgin_FechaRelatorioGerencial
                End If
            End If
        End If
        xTipoTEF = "LeituraX"
        If xSaiDoLoop = True Then
            Exit Do
        End If
    Loop
    
    'Desbloqueia Teclado
    If gTipoDocumentoFiscal <> "NFCe" Then
        Call Bematech_FI_FinalizaModoTEF
    End If
    
    Exit Function
FileError:
    'Desbloqueia Teclado
    If gTipoDocumentoFiscal <> "NFCe" Then
        Call Bematech_FI_FinalizaModoTEF
    End If
    Call CriaLogTEF(Date & " " & Time & " TestaImprimeTefADM: " & " - ErroNúmero: " & Err & " - ErroTexto: " & Error)
End Function
Function TestaSolicitacaoADM() As Boolean
    Dim i As Integer
    Dim xCampo009 As String
    Dim xCampo028 As Long
    Dim xNomeArquivoResp001 As String 'Arquivo de Resposta
    Dim xExisteResposta As Boolean
    Dim xUsarCampo715 As Boolean
    
    On Error GoTo FileError
    xCampo009 = ""
    xCampo028 = 0
    xUsarCampo715 = False
    TestaSolicitacaoADM = False
    lCampo001 = False
    
'    If UCase(gBandeira) = "TECBAN" Then
'        xNomeArquivoResp001 = "C:\TEF_DISC\RESP\IntPos.001"
'    ElseIf UCase(gBandeira) = "TCSMART" Then
'        xNomeArquivoResp001 = "C:\TCS\RX\IntTcs.001"
'    ElseIf UCase(gBandeira) = "SMARTEF" Then
'        xNomeArquivoResp001 = "C:\SMARTEF\RESP\IntPos.001"
'    ElseIf UCase(gBandeira) = "SUPERTEF" Then
'        xNomeArquivoResp001 = "C:\SUPERTEF\RESP\IntPos.001"
'    ElseIf UCase(gBandeira) = "HIPERTEF" Then
'        xNomeArquivoResp001 = "C:\HiperTEF\RESP\IntPos.001"
'    ElseIf UCase(gBandeira) = "PAGCARC" Then
'        xNomeArquivoResp001 = "C:\CardTech_NEUS\RESP\IntPos.001"
'    Else
'        xNomeArquivoResp001 = "C:\TEF_DIAL\RESP\IntPos.001"
'    End If
    If UCase(gBandeira) = "TCSMART" Then
        xNomeArquivoResp001 = gDiretorioResp & "IntTcs.001"
    Else
        xNomeArquivoResp001 = gDiretorioResp & "IntPos.001"
    End If
    
    rtxt_mensagem.Text = "Aguardando Operação" & Chr(10) & "Administrativa!"
    'Call CriaLogTEF("TestaSolicitacaoADM: Mensagem 10: " & rtxt_mensagem.Text & " às: " & Time)
    rtxt_mensagem.Visible = True
    Do Until gArqTxt.FileExists(xNomeArquivoResp001)
        DoEvents
    Loop
    rtxt_mensagem.Visible = False
    
    
    'Mostra mensagem 30
    rtxt_mensagem.Enabled = False
    If gArqTxt.FileExists(xNomeArquivoResp001) Then
        Call CriaLogTefEspecial("TestaSolicitacaoADM - 001 Arquivo de retorno encontrado. xNomeArquivoResp001: " & xNomeArquivoResp001 & " - gNumeroControleSolicitacao: " & gNumeroControleSolicitacao & " - gBandeira: " & gBandeira)
        'BUG:XP DÁ UM TEMPO PARA LIBERAR ARQUIVO PARA LEITURA
        lHoraInicial = Time
        Do Until DateDiff("s", lHoraInicial, Time) >= 5
        Loop
        Set gArquivo = gArqTxt.OpenTextFile(xNomeArquivoResp001, ForReading)
        lMensagem = ""
        i = 0
        Do Until gArquivo.AtEndOfStream
            lString = gArquivo.ReadLine
            Call CriaLogTefEspecial("TestaSolicitacaoADM - 002 Dados da linha de retorno. lString: " & lString)
            If UCase(gBandeira) = "TCSMART" Then
                If Mid(lString, 1, 7) = "501-000" Then
                    If Mid(lString, 11, 10) = gNumeroControleSolicitacao Then
                        lCampo001 = True
                        xCampo009 = "0"
                    End If
                End If
                If Mid(lString, 1, 3) = "514" Then
                    i = i + 1
                    If i > 1 Then
                        lMensagem = lMensagem & Chr(10)
                    End If
                    xCampo028 = Mid(lString, 11, Len(lString) - 10)
                    lMensagem = lMensagem & Mid(lString, 11, Len(lString) - 10) & " Linhas à Imprimir"
                End If
            Else
                If Mid(lString, 1, 7) = "001-000" Then
                    Call CriaLogTefEspecial("TestaSolicitacaoADM - 003 Testa campo 001-000. Esperando que gNumeroControleSolicitacao seja: " & gNumeroControleSolicitacao)
                    If Mid(lString, 11, 10) = gNumeroControleSolicitacao Then
                        Call CriaLogTefEspecial("TestaSolicitacaoADM - 004 Campo lCampo001 passa a ser verdadeiro. gNumeroControleSolicitacao: " & gNumeroControleSolicitacao)
                        lCampo001 = True
                    End If
                End If
                If Mid(lString, 1, 7) = "009-000" Then
                    xCampo009 = Mid(lString, 11, Len(lString) - 10)
                End If
                If Mid(lString, 1, 7) = "028-000" Then
                    xCampo028 = Mid(lString, 11, Len(lString) - 10)
                End If
                If Mid(lString, 1, 7) = "714-000" Then
                    Call CriaLogTefEspecial("TestaSolicitacaoADM - 101 Testa campo 714-000. E verifica se xCampo028 = 0. xCampo028=" & xCampo028)
                    If xCampo028 = 0 Then
                        Call CriaLogTefEspecial("TestaSolicitacaoADM - 102 Redefine xCampo028 e xUsarCampo715 = True.")
                        xCampo028 = Mid(lString, 11, Len(lString) - 10)
                        xUsarCampo715 = True
                    End If
                End If
                If Mid(lString, 1, 3) = "030" Then
                    i = i + 1
                    If i > 1 Then
                        lMensagem = lMensagem & Chr(10)
                    End If
                    lMensagem = lMensagem & Mid(lString, 11, Len(lString) - 10)
                End If
            End If
        Loop
        rtxt_mensagem.Text = lMensagem
        'Call CriaLogTEF("TestaSolicitacaoADM: Mensagem 11: " & rtxt_mensagem.Text & " às: " & Time)
        rtxt_mensagem.Visible = True
        If xCampo009 = "0" Or xCampo009 = "P1" Then
            'lHoraInicial = Time
            '''Do Until DateDiff("s", lHoraInicial, Time) >= 10
            '''    DoEvents
            '''Loop
        End If
        '''rtxt_mensagem.Visible = False
        gArquivo.Close
        
        Call CriaLogTefEspecial("TestaSolicitacaoADM - 004 Testa campos. lCampo001: " & lCampo001 & " - xCampo009: " & xCampo009 & " - xCampo028: " & xCampo028)
        If lCampo001 Then
            If xCampo009 <> "0" And xCampo028 = 0 Then
                MsgBox lMensagem, vbInformation, "Mensagem ao Operador"
            End If
            If gTefString = "SolicitacaoConsultaCH" And xCampo009 = "0" And xCampo028 = 0 Then
                MsgBox lMensagem, vbInformation, "Mensagem ao Operador"
            End If
            If xCampo009 = "0" Or xCampo009 = "00" Then
                If UCase(gBandeira) = "TCSMART" Then
                Else
                    'Fazer backup do "IntPos.001" para c:\Backup_Sgp
                    Call CopiarArquivo(xNomeArquivoResp001, "C:\Backup_SGP\IntPos.001")
                End If
                If xCampo028 = 0 Then
                    Call ExcluirArquivo(xNomeArquivoResp001)
                Else
                    TestaSolicitacaoADM = True
                    'ImprimeTef
                End If
            Else
                'MsgBox lMensagem, vbInformation, "Mensagem de retorno"
                Call ExcluirArquivo(xNomeArquivoResp001)
            End If
        Else
            MsgBox "Numero de controle de solicitação está diferente", vbInformation, "TestaSolicitacaoADM"
        End If
    Else
        Call CriaLogTefEspecial("SolicitacaoADM - 077 Arquivo de retorno INEXISTENTE. xNomeArquivoResp001: " & xNomeArquivoResp001)
        Call CriaLogTefEspecial("TestaSolicitacaoADM - 088 Arquivo de retorno INEXISTENTE. xNomeArquivoResp001: " & xNomeArquivoResp001)
        MsgBox "O arquivo IntPos.001 foi deletado de forma não identificada!", vbInformation, "Mensagem Padrão"
    End If
    Exit Function
FileError:
    Call CriaLogTefEspecial("TestaSolicitacaoADM - 099 Erro Número: " & Err & " - ErroTexto: " & Error)
    Call CriaLogTEF(Date & " " & Time & " Erro TestaSolicitacaoADM: " & " - ErroNúmero: " & Err & " - ErroTexto: " & Error)
    MsgBox "ERRO não identificado!", vbInformation, "Rotina: SolicitacaoDeCompra"
    Exit Function
End Function
Function TestaAlteraPrecoTCS() As Boolean
    Dim i As Integer
    Dim xNomeArquivoResp001 As String
    
On Error GoTo FileError
    
    TestaAlteraPrecoTCS = False
    xNomeArquivoResp001 = gDiretorioResp & "IntTcs.RET"
    Do Until gArqTxt.FileExists(xNomeArquivoResp001)
        DoEvents
    Loop
    
    If gArqTxt.FileExists(xNomeArquivoResp001) Then
        Set gArquivo = gArqTxt.OpenTextFile(xNomeArquivoResp001, ForReading)
        lString = gArquivo.ReadLine
        If Mid(lString, 1, 7) = "517-000" Then
            If Mid(lString, 11, 3) = "000" Then
                TestaAlteraPrecoTCS = True
            Else
                lString = gArquivo.ReadLine
                lString = Mid(lString, 12, Len(lString) - 12)
                MsgBox lString, vbInformation, "TestaAlteraPrecoTCS"
            End If
        End If
        gArquivo.Close
        Call ExcluirArquivo(xNomeArquivoResp001)
    End If
    Exit Function
    
FileError:
    Call CriaLogTEF(Date & " " & Time & " Erro TestaAlteraPrecoTCS: - ErroTexto: " & Error)
    Exit Function
End Function
Function TestaPendencia() As Boolean
    Dim xNomeArquivoResp As String 'Arquivo de Resposta
    TestaPendencia = False
    rtxt_mensagem.Text = "Testando Pendência!"
    'Call CriaLogTEF("TestaPendencia: Mensagem 12: " & rtxt_mensagem.Text & " às: " & Time)
    rtxt_mensagem.Visible = True
    xNomeArquivoResp = "C:\TEF_DIAL\RESP\IntPos.001"
    If gArqTxt.FileExists(xNomeArquivoResp) Then
        TestaPendencia = True
    End If
    xNomeArquivoResp = "C:\TEF_DISC\RESP\IntPos.001"
    If gArqTxt.FileExists(xNomeArquivoResp) Then
        TestaPendencia = True
    End If
    xNomeArquivoResp = "C:\TCS\RX\IntTcs.001"
    If gArqTxt.FileExists(xNomeArquivoResp) Then
        TestaPendencia = True
    End If
    xNomeArquivoResp = "C:\SUPERTEF\RESP\IntPos.001"
    If gArqTxt.FileExists(xNomeArquivoResp) Then
        TestaPendencia = True
    End If
    xNomeArquivoResp = "C:\HiperTEF\RESP\IntPos.001"
    If gArqTxt.FileExists(xNomeArquivoResp) Then
        TestaPendencia = True
    End If
    xNomeArquivoResp = "C:\CardTech_NEUS\RESP\IntPos.001"
    If gArqTxt.FileExists(xNomeArquivoResp) Then
        TestaPendencia = True
    End If
    xNomeArquivoResp = "C:\TefCerrado\RESP\IntPos.001"
    If gArqTxt.FileExists(xNomeArquivoResp) Then
        TestaPendencia = True
    End If
    xNomeArquivoResp = "C:\Tef_NEUS\RESP\IntPos.001"
    If gArqTxt.FileExists(xNomeArquivoResp) Then
        TestaPendencia = True
    End If
    xNomeArquivoResp = "C:\GetNet\RESP\IntPos.001"
    If gArqTxt.FileExists(xNomeArquivoResp) Then
        TestaPendencia = True
    End If
End Function
Function TestaSolicitacao() As Boolean
    Dim xNomeArquivoResp As String 'Arquivo de Resposta
    Dim xFase As Integer
    
    On Error GoTo FileError
    
    TestaSolicitacao = False
    xFase = 0
'    If UCase(gBandeira) = "TECBAN" Then
'        xNomeArquivoResp = "C:\TEF_DISC\RESP\IntPos.001"
'    ElseIf UCase(gBandeira) = "TCSMART" Then
'        xNomeArquivoResp = "C:\TCS\RX\IntTcs.001"
'    ElseIf UCase(gBandeira) = "SMARTEF" Then
'        xNomeArquivoResp = "C:\SMARTEF\RESP\IntPos.001"
'    ElseIf UCase(gBandeira) = "SUPERTEF" Then
'        xNomeArquivoResp = "C:\SUPERTEF\RESP\IntPos.001"
'    ElseIf UCase(gBandeira) = "HIPERTEF" Then
'        xNomeArquivoResp = "C:\HiperTEF\RESP\IntPos.001"
'    ElseIf UCase(gBandeira) = "PAGCARD" Then
'        xNomeArquivoResp = "C:\CardTech_NEUS\RESP\IntPos.001"
'    Else
'        xNomeArquivoResp = "C:\TEF_DIAL\RESP\IntPos.001"
'    End If
    If UCase(gBandeira) = "TCSMART" Then
        xNomeArquivoResp = gDiretorioResp & "IntTcs.001"
    Else
        xNomeArquivoResp = gDiretorioResp & "IntPos.001"
    End If
    xFase = 1
    
    rtxt_mensagem.Text = "Aguardando Operação!"
    'Call CriaLogTEF("TestaSolicitacao: Mensagem 13: " & rtxt_mensagem.Text & " às: " & Time)
    rtxt_mensagem.Visible = True
    Do Until gArqTxt.FileExists(xNomeArquivoResp)
        DoEvents
    Loop
    
    If UCase(gBandeira) = "TCSMART" Then
    Else
        Call LoopNumeroControleSolicitacao(xNomeArquivoResp)
    End If

    If CarregaMensagemTEF(True, False, False) Then
        If lCampo001 Then
            If lCampo009 <> "0" And lCampo028 = 0 Then
                MsgBox lMensagem, vbInformation, "Mensagem ao Operador"
            End If
            If gConsultaCheque Then
                If lCampo009 = "0" Or lCampo009 = "P1" Then
                    TestaSolicitacao = True
                    'Fazer backup do "IntPos.001" para c:\Backup_Sgp
                    xFase = 3
                    Call CopiarArquivo(xNomeArquivoResp, "C:\Backup_SGP\IntPos.001")
                End If
                'no caso consulta de cheque redecard nao escluir o arquivo abaixo
                'xFase = 4
                'Call ExcluirArquivo(xNomeArquivoResp)
                'xFase = 5
            Else
                If lCampo009 = "0" Then
                    If lCampo028 = 0 Then
                        xFase = 6
                        Call ExcluirArquivo(xNomeArquivoResp)
                        xFase = 7
                    Else
                        TestaSolicitacao = True
                        'Fazer backup do "IntPos.001" para c:\Backup_Sgp
                        xFase = 8
                        If UCase(gBandeira) = "TCSMART" Then
                        Else
                            Call CopiarArquivo(xNomeArquivoResp, "C:\Backup_SGP\IntPos.001")
                        End If
                        'ImprimeTef
                    End If
                Else
                    'MsgBox lMensagem, vbInformation, "Mensagem de retorno"
                    xFase = 9
                    Call ExcluirArquivo(xNomeArquivoResp)
                    xFase = 10
                End If
            End If
        Else
            MsgBox "Numero de controle de solicitação está diferente", vbInformation, "TestaSolicitacao"
        End If
    Else
        MsgBox lMensagem, vbInformation, "Mensagem para o Operador!"
        xFase = 11
        Call ExcluirArquivo(xNomeArquivoResp)
        xFase = 12
    End If
    Exit Function

FileError:
    Call CriaLogTEF(Date & " " & Time & " Erro TestaSolicitacao: " & " - ErroNúmero: " & Err & " - ErroTexto: " & Error)
    Call CriaLogTEF(Date & " " & Time & " Erro TestaSolicitacao: Fase: " & xFase)
    MsgBox "ERRO não identificado!", vbInformation, "Rotina: TestaSolicitacao"
    Exit Function

End Function
Private Sub Form_Unload(Cancel As Integer)
    Set lControleSolicitacao = Nothing
End Sub
Private Sub txt_nsu_GotFocus()
    txt_nsu.SelStart = 0
    txt_nsu.SelLength = Len(txt_nsu.Text)
End Sub
Private Sub txt_nsu_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_valor_GotFocus()
    txt_valor.SelStart = 0
    txt_valor.SelLength = Len(txt_valor.Text)
End Sub
Private Sub txt_valor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_nsu.SetFocus
    End If
    'Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_valor_LostFocus()
    If txt_valor.Text <> "" Then
        txt_valor.Text = Format(txt_valor.Text, "###,##0.00")
    End If
End Sub
