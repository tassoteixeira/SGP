VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form analizador_tef 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "T.E.F."
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6210
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   6210
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox rtxt_mensagem 
      Height          =   3375
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   5953
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      RightMargin     =   9,99999e5
      TextRTF         =   $"analizador_tef.frx":0000
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
End
Attribute VB_Name = "analizador_tef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lArqTxt As New FileSystemObject
Dim lString As String
Dim lNumeroArquivo As Integer
Dim lCampo001000Igual As Boolean
Dim lNumeroCupom As Long
Dim lValorRecebido As String
Dim lMensagem As String
Dim lMensagem29(0 To 60) As String
Dim lQtdMensagem29 As Integer
Dim lHoraInicial As Date
Dim BemaRetorno As Integer
Private Sub Form_Activate()
    If RetiraGTefString(1) = "GerenciadorPadraoAtivo" Then
        rtxt_mensagem = "Verifica se o" & Chr(10) & "Gerenciador Padrão" & Chr(10) & "Está ativo!"
        rtxt_mensagem.Visible = True
        DoEvents
        gTefResposta = GerenciadorPadraoAtivo
        rtxt_mensagem.Visible = False
    ElseIf RetiraGTefString(1) = "SolicitacaoADM" Then
        rtxt_mensagem = "Solicitação de Funções" & Chr(10) & "Administrativas!"
        rtxt_mensagem.Visible = True
        DoEvents
        gTefResposta = SolicitacaoADM
        rtxt_mensagem.Visible = False
    ElseIf RetiraGTefString(1) = "SolicitacaoDeCompra" Then
        rtxt_mensagem = "Solicitação de Compra!"
        rtxt_mensagem.Visible = True
        DoEvents
        lNumeroCupom = RetiraGTefString(2)
        lValorRecebido = RetiraGTefString(3)
        gTefResposta = SolicitacaoDeCompra
        rtxt_mensagem.Visible = False
    ElseIf RetiraGTefString(1) = "TestaSolicitacaoADM" Then
        gTefResposta = TestaSolicitacaoADM
    ElseIf RetiraGTefString(1) = "TestaSolicitacaoDeCompra" Then
        lNumeroCupom = RetiraGTefString(2)
        lValorRecebido = RetiraGTefString(3)
        gTefResposta = TestaSolicitacaoDeCompra
    ElseIf RetiraGTefString(1) = "ImprimeTefADM" Then
        rtxt_mensagem = "Aguarde!" & Chr(10) & Chr(10) & "Imprimindo TEF" & Chr(10) & "Administrativo."
        rtxt_mensagem.Visible = True
        DoEvents
        gTefResposta = TestaImprimeTefADM("LeituraX")
        rtxt_mensagem.Visible = False
    ElseIf RetiraGTefString(1) = "ImprimeTEF" Then
        rtxt_mensagem = "Aguarde!" & Chr(10) & Chr(10) & "Imprimindo TEF."
        rtxt_mensagem.Visible = True
        DoEvents
        lNumeroCupom = RetiraGTefString(2)
        lValorRecebido = RetiraGTefString(3)
        gTefResposta = TestaImprimeTEF("Vinculado")
        rtxt_mensagem.Visible = False
    ElseIf RetiraGTefString(1) = "CNF" Then
        rtxt_mensagem = "Aguarde!" & Chr(10) & Chr(10) & "Confimando TEF."
        rtxt_mensagem.Visible = True
        DoEvents
        'lNumeroCupom = RetiraGTefString(2)
        'lValorRecebido = RetiraGTefString(3)
        gTefResposta = CNF
        rtxt_mensagem.Visible = False
    ElseIf RetiraGTefString(1) = "NCN" Then
        rtxt_mensagem = "Aguarde!" & Chr(10) & Chr(10) & "Cancelando TEF."
        rtxt_mensagem.Visible = True
        DoEvents
        'lNumeroCupom = RetiraGTefString(2)
        'lValorRecebido = RetiraGTefString(3)
        gTefResposta = NCN
        rtxt_mensagem.Visible = False
    End If
    Call DisabelCtrlAltDel(False)
    Unload Me
End Sub
Private Sub Form_Load()
    CentraForm Me
    gTefResposta = False
    rtxt_mensagem.Visible = False
    Call DisabelCtrlAltDel(True)
End Sub
Function ImprimeTEF(ByVal xTipoImpressao As String) As Boolean
    Dim i As Integer
    Dim xVias As Integer
    Dim xValorRecebido
    ImprimeTEF = False
    
    
    xValorRecebido = Mid(Format(lValorRecebido, "000000000000.00"), 1, 12) & Mid(Format(lValorRecebido, "000000000000.00"), 14, 2)
    
    
    'Abre Relatorio Gerencial
    If xTipoImpressao = "LeituraX" Then
        BemaRetorno = Bematech_FI_RelatorioGerencial(" ")
    Else
        BemaRetorno = Bematech_FI_AbreComprovanteNaoFiscalVinculado("Cartao credito  ", xValorRecebido, CStr(Format(lNumeroCupom, "000000")))
    End If
    'Verifica se Imprimiu
    If BemaRetorno = 1 Then
        For xVias = 1 To gQtdViasTEF
            For i = 0 To lQtdMensagem29
                'Imprime Texto do TEF
                If xTipoImpressao = "LeituraX" Then
                    BemaRetorno = Bematech_FI_RelatorioGerencial(lMensagem29(i))
                Else
                    BemaRetorno = Bematech_FI_UsaComprovanteNaoFiscalVinculado(lMensagem29(i))
                End If
                'Verifica se Não Imprimiu
                If BemaRetorno <> 1 Then
                    Exit Function
                End If
                If xVias = 2 And i = 4 Then
                    'Pausa de 5 Segundos Entre a Primeira e Segunda Via
                    'para Cortar o Papel
                    rtxt_mensagem = "Primeira Via Já Está Impressa." & Chr(10) & "Recorte Agora!"
                    rtxt_mensagem.Visible = True
                    DoEvents
                    lHoraInicial = Time
                    Me.Caption = "Mensagem para o Operador"
                    Do Until DateDiff("s", lHoraInicial, Time) >= 5
                        DoEvents
                    Loop
                    rtxt_mensagem = "Aguarde!" & Chr(10) & Chr(10) & "Imprimindo TEF"
                    rtxt_mensagem.Visible = True
                    DoEvents
                End If
            Next
            If gQtdViasTEF = 2 Then
                'Imprime 2 linhas em branco
                If xTipoImpressao = "LeituraX" Then
                    BemaRetorno = Bematech_FI_RelatorioGerencial(Space(96))
                Else
                    BemaRetorno = Bematech_FI_UsaComprovanteNaoFiscalVinculado(Space(96))
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
        BemaRetorno = Bematech_FI_FechaRelatorioGerencial
    Else
        BemaRetorno = Bematech_FI_FechaComprovanteNaoFiscalVinculado
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
    Dim xVias As Integer
    ImprimeTefADM = False
    
    'Abre Relatorio Gerencial
    BemaRetorno = Bematech_FI_RelatorioGerencial(" ")
    'Verifica se Imprimiu
    If BemaRetorno = 1 Then
        For xVias = 1 To gQtdViasTEF
            For i = 0 To lQtdMensagem29
                'Imprime Texto do TEF
                BemaRetorno = Bematech_FI_RelatorioGerencial(lMensagem29(i))
                'Verifica se Não Imprimiu
                If BemaRetorno <> 1 Then
                    Exit Function
                End If
                If xVias = 2 And i = 4 Then
                    'Pausa de 5 Segundos Entre a Primeira e Segunda Via
                    'para Cortar o Papel
                    rtxt_mensagem = "Primeira Via Já Está Impressa." & Chr(10) & "Recorte Agora!"
                    rtxt_mensagem.Visible = True
                    DoEvents
                    lHoraInicial = Time
                    Me.Caption = "Mensagem para o Operador"
                    Do Until DateDiff("s", lHoraInicial, Time) >= 5
                        DoEvents
                    Loop
                    rtxt_mensagem = "Aguarde!" & Chr(10) & Chr(10) & "Imprimindo TEF" & Chr(10) & "Administrativo."
                    rtxt_mensagem.Visible = True
                    DoEvents
                End If
            Next
            If gQtdViasTEF = 2 Then
                'Imprime 2 linhas em branco
                BemaRetorno = Bematech_FI_RelatorioGerencial(Space(96))
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
    BemaRetorno = Bematech_FI_FechaRelatorioGerencial
    'Verifica se Imprimiu
    If BemaRetorno = 1 Then
        ImprimeTefADM = True
    Else
        Exit Function
    End If
End Function
Function AtivaGerenciadorPadrao() As Boolean
    Dim RetVal As Long
    AtivaGerenciadorPadrao = False
    RetVal = Shell("C:\tef_dial\tef_dial.exe", vbMinimizedNoFocus)
End Function
Function CarregaMensagemTEF() As Boolean
    Dim i As Integer
    Dim i2 As Integer
    CarregaMensagemTEF = False
    If lArqTxt.FileExists("C:\TEF_DIAL\RESP\IntPos.001") Then
        lNumeroArquivo = FreeFile
        Open "C:\TEF_DIAL\RESP\IntPos.001" For Input As #lNumeroArquivo
        lMensagem = ""
        i = 0
        lQtdMensagem29 = -1
        Do Until EOF(lNumeroArquivo)
            Input #lNumeroArquivo, lString
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
        Loop
        rtxt_mensagem = lMensagem
        rtxt_mensagem.Visible = True
        lHoraInicial = Time
        Me.Caption = "Mensagem para o Operador"
        Do Until DateDiff("s", lHoraInicial, Time) >= 5
            DoEvents
        Loop
        Me.Caption = "T.E.F."
        rtxt_mensagem.Visible = False
        Close #lNumeroArquivo
    Else
        MsgBox "O arquivo IntPos.001 foi deletado de forma não identificada!", vbInformation, "Mensagem Padrão"
        Exit Function
    End If
    If lQtdMensagem29 <> -1 Then
        CarregaMensagemTEF = True
    End If
End Function
Function CNF() As Boolean
    Dim xString As String
    Dim xValorRecebido As String
    Dim xNomeRede As String
    Dim xNumeroTransacao As String
    Dim xFinalizacao As String
    Dim i As Integer
    Dim i2 As Integer
    CNF = False
    If lArqTxt.FileExists("C:\TEF_DIAL\RESP\IntPos.001") Then
        lNumeroArquivo = FreeFile
        Open "C:\TEF_DIAL\RESP\IntPos.001" For Input As #lNumeroArquivo
        lMensagem = ""
        i = 0
        lQtdMensagem29 = -1
        Do Until EOF(lNumeroArquivo)
            Input #lNumeroArquivo, lString
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
        Close #lNumeroArquivo
    Else
        MsgBox "O arquivo IntPos.001 foi deletado de forma não identificada!", vbInformation, "Mensagem Padrão"
        Exit Function
    End If
    Open "C:\TEF_DIAL\REQ\IntPos.tmp" For Output As #lNumeroArquivo
    Print #lNumeroArquivo, "000-000 = CNF"
    Print #lNumeroArquivo, "001-000 = " & Format(gNumeroControleSolicitacao, "0000000000")
    Print #lNumeroArquivo, "010-000 = " & xNomeRede
    Print #lNumeroArquivo, "012-000 = " & xNumeroTransacao
    Print #lNumeroArquivo, "027-000 = " & xFinalizacao
    Print #lNumeroArquivo, "999-999 = 0"
    Close #lNumeroArquivo
    If lArqTxt.FileExists("C:\TEF_DIAL\REQ\IntPos.tmp") Then
        lArqTxt.MoveFile ("C:\TEF_DIAL\REQ\IntPos.tmp"), ("C:\TEF_DIAL\REQ\IntPos.001")
    End If
    
    lHoraInicial = Time
    Do Until DateDiff("s", lHoraInicial, Time) >= 7
        If lArqTxt.FileExists("C:\TEF_DIAL\RESP\IntPos.STS") Then
            Exit Do
        End If
        DoEvents
    Loop
    If lArqTxt.FileExists("C:\TEF_DIAL\RESP\IntPos.STS") Then
        lNumeroArquivo = FreeFile
        Open "C:\TEF_DIAL\RESP\IntPos.STS" For Input As #lNumeroArquivo
        i = 0
        lCampo001000Igual = False
        Do Until EOF(lNumeroArquivo)
            Input #lNumeroArquivo, lString
            i = i + 1
            If i = 1 Then
                If lString <> "000-000 = CNF" Then
                    Exit Do
                End If
            End If
            If Mid(lString, 1, 7) = "001-000" Then
                If Mid(lString, 11, 10) = gNumeroControleSolicitacao Then
                    lCampo001000Igual = True
                End If
            End If
        Loop
        Close #lNumeroArquivo
        lArqTxt.DeleteFile ("C:\TEF_DIAL\RESP\IntPos.STS")
        If lCampo001000Igual Then
            lArqTxt.DeleteFile ("C:\TEF_DIAL\RESP\IntPos.001")
            CNF = True
        Else
            MsgBox "Numero de controle de solicitação está diferente", vbInformation, "????"
        End If
    Else
        MsgBox "TEF Não Está Ativo!", vbInformation, "Mensagem Padrão"
    End If
End Function
Function GerenciadorPadraoAtivo() As Boolean
    Dim i As Integer
    On Error GoTo FileError
    GerenciadorPadraoAtivo = False
    lCampo001000Igual = False
    'ATV
    'Se existir arquivos, deleta.
    If lArqTxt.FileExists("C:\TEF_DIAL\REQ\IntPos.tmp") Then
        lArqTxt.DeleteFile ("C:\TEF_DIAL\REQ\IntPos.tmp")
    End If
    If lArqTxt.FileExists("C:\TEF_DIAL\REQ\IntPos.001") Then
        lArqTxt.DeleteFile ("C:\TEF_DIAL\REQ\IntPos.001")
    End If
    'Gera Arquivo
    lNumeroArquivo = FreeFile
    Open "C:\TEF_DIAL\REQ\IntPos.tmp" For Output As #lNumeroArquivo
    Print #lNumeroArquivo, "000-000 = ATV"
    'Print #lNumeroArquivo, "000-000 = ADM"
    Print #lNumeroArquivo, "001-000 = " & Format(gNumeroControleSolicitacao, "0000000000")
    Print #lNumeroArquivo, "999-999 = 0"
    Close #lNumeroArquivo
    If lArqTxt.FileExists("C:\TEF_DIAL\REQ\IntPos.tmp") Then
        lArqTxt.MoveFile ("C:\TEF_DIAL\REQ\IntPos.tmp"), ("C:\TEF_DIAL\REQ\IntPos.001")
    End If
    'Aguarda 7 segundos
    lHoraInicial = Time
    Do Until DateDiff("s", lHoraInicial, Time) >= 7
        If lArqTxt.FileExists("C:\TEF_DIAL\RESP\IntPos.STS") Then
            Exit Do
        End If
        DoEvents
    Loop
    
    If Not lArqTxt.FileExists("C:\TEF_DIAL\RESP\IntPos.STS") Then
        MsgBox "Gerenciador Padrão não está ativo, e será ativado automaticamente!", vbInformation, "Mensagem Padrão"
        Call AtivaGerenciadorPadrao
        'Aguarda 7 segundos
        lHoraInicial = Time
        Do Until DateDiff("s", lHoraInicial, Time) >= 7
            If lArqTxt.FileExists("C:\TEF_DIAL\RESP\IntPos.STS") Then
                Exit Do
            End If
            DoEvents
        Loop
    End If
    
    If lArqTxt.FileExists("C:\TEF_DIAL\RESP\IntPos.STS") Then
        lNumeroArquivo = FreeFile
        Open "C:\TEF_DIAL\RESP\IntPos.STS" For Input As #lNumeroArquivo
        i = 0
        Do Until EOF(lNumeroArquivo)
            Input #lNumeroArquivo, lString
            i = i + 1
            If i = 1 Then
                If lString <> "000-000 = ATV" Then
                    Exit Do
                End If
            End If
            'Verifica se o número do controle da solicitação é igual
            If Mid(lString, 1, 7) = "001-000" Then
                If Mid(lString, 11, 10) = gNumeroControleSolicitacao Then
                    lCampo001000Igual = True
                End If
            End If
        Loop
        Close #lNumeroArquivo
        lArqTxt.DeleteFile ("C:\TEF_DIAL\RESP\IntPos.STS")
        If lCampo001000Igual Then
            GerenciadorPadraoAtivo = True
        Else
            MsgBox "Numero de controle de solicitação está diferente", vbInformation, "????"
        End If
    Else
        MsgBox "Gerenciador Padrão não está ativo!", vbInformation, "Mensagem Padrão"
    End If
    Exit Function
FileError:
    If Err = 76 Then
        MsgBox "Gerenciador Padrão não está instalado neste computador!", vbInformation, "Transação com TEF não aceita!"
        Exit Function
    End If
    MsgBox "ERRO não identificado!", vbInformation, "Rotina: GerenciadorPadraoAtivo"
    Exit Function
End Function
Function NCN() As Boolean
    Dim xString As String
    Dim xValorRecebido As String
    Dim xNomeRede As String
    Dim xNumeroTransacao As String
    Dim xFinalizacao As String
    Dim xValor As String
    Dim xCampo10 As String
    Dim xCampo12 As String
    Dim i As Integer
    Dim i2 As Integer
    NCN = False
    If lArqTxt.FileExists("C:\TEF_DIAL\RESP\IntPos.001") Then
        lNumeroArquivo = FreeFile
        Open "C:\TEF_DIAL\RESP\IntPos.001" For Input As #lNumeroArquivo
        lMensagem = ""
        i = 0
        lQtdMensagem29 = -1
        Do Until EOF(lNumeroArquivo)
            Input #lNumeroArquivo, lString
            'Guarda Texto da Finalizacao
            If Mid(lString, 1, 3) = "003" Then
                i2 = Len(lString)
                If Len(lString) > 11 Then
                    xValor = Mid(lString, 11, i2 - 10)
                End If
            End If
            If Mid(lString, 1, 3) = "010" Then
                i2 = Len(lString)
                If Len(lString) > 11 Then
                    xNomeRede = Mid(lString, 11, i2 - 10)
                End If
            End If
            If Mid(lString, 1, 3) = "012" Then
                i2 = Len(lString)
                xNumeroTransacao = Mid(lString, 11, i2 - 10)
            End If
            If Mid(lString, 1, 3) = "027" Then
                i2 = Len(lString)
                If Len(lString) > 11 Then
                    xFinalizacao = Mid(lString, 11, i2 - 10)
                End If
                Exit Do
            End If
        Loop
        Close #lNumeroArquivo
        lArqTxt.DeleteFile ("C:\TEF_DIAL\RESP\IntPos.001")
    Else
        MsgBox "O arquivo IntPos.001 foi deletado de forma não identificada!", vbInformation, "Mensagem Padrão"
        Exit Function
    End If
    
    xString = ""
    If Val(xValor) > 0 Then
        i = Len(xValor)
        xString = Chr(10) & Chr(10) & "Valor: " & Format(Mid(xValor, 1, i - 2) & "," & Mid(xValor, i - 1, 2), "###,###,##0.00")
    End If
    Open "C:\TEF_DIAL\REQ\IntPos.tmp" For Output As #lNumeroArquivo
    Print #lNumeroArquivo, "000-000 = NCN"
    Print #lNumeroArquivo, "001-000 = " & Format(gNumeroControleSolicitacao, "0000000000")
    Print #lNumeroArquivo, "010-000 = " & xNomeRede
    Print #lNumeroArquivo, "012-000 = " & xNumeroTransacao
    Print #lNumeroArquivo, "027-000 = " & xFinalizacao
    Print #lNumeroArquivo, "999-999 = 0"
    Close #lNumeroArquivo
    If lArqTxt.FileExists("C:\TEF_DIAL\REQ\IntPos.tmp") Then
        lArqTxt.MoveFile ("C:\TEF_DIAL\REQ\IntPos.tmp"), ("C:\TEF_DIAL\REQ\IntPos.001")
    End If
    
    lHoraInicial = Time
    Do Until DateDiff("s", lHoraInicial, Time) >= 7
        If lArqTxt.FileExists("C:\TEF_DIAL\RESP\IntPos.STS") Then
            Exit Do
        End If
        DoEvents
    Loop
    If lArqTxt.FileExists("C:\TEF_DIAL\RESP\IntPos.STS") Then
        lNumeroArquivo = FreeFile
        Open "C:\TEF_DIAL\RESP\IntPos.STS" For Input As #lNumeroArquivo
        i = 0
        lCampo001000Igual = False
        Do Until EOF(lNumeroArquivo)
            Input #lNumeroArquivo, lString
            i = i + 1
            If i = 1 Then
                If lString <> "000-000 = NCN" Then
                    Exit Do
                End If
            End If
            If Mid(lString, 1, 7) = "001-000" Then
                If Mid(lString, 11, 10) = gNumeroControleSolicitacao Then
                    lCampo001000Igual = True
                End If
            End If
        Loop
        Close #lNumeroArquivo
        lArqTxt.DeleteFile ("C:\TEF_DIAL\RESP\IntPos.STS")
        If lCampo001000Igual Then
            NCN = True
            MsgBox "Última Transação TEF foi Cancelada" & Chr(10) & Chr(10) & "Rede: " & xNomeRede & Chr(10) & Chr(10) & "NSU: " & Format(xNumeroTransacao, "###########0") & xString, vbInformation, "Última transação TEF foi cancelada."
        Else
            MsgBox "Numero de controle de solicitação está diferente", vbInformation, "????"
        End If
    Else
        MsgBox "TEF Não Está Ativo!", vbInformation, "Mensagem Padrão"
    End If
End Function
Function SolicitacaoADM() As Boolean
    Dim i As Integer
    On Error GoTo FileError
    SolicitacaoADM = False
    lCampo001000Igual = False
    'ADM
    'Se existir arquivos, deleta.
    If lArqTxt.FileExists("C:\TEF_DIAL\REQ\IntPos.tmp") Then
        lArqTxt.DeleteFile ("C:\TEF_DIAL\REQ\IntPos.tmp")
    End If
    If lArqTxt.FileExists("C:\TEF_DIAL\REQ\IntPos.001") Then
        lArqTxt.DeleteFile ("C:\TEF_DIAL\REQ\IntPos.001")
    End If
    'Gera Arquivo
    lNumeroArquivo = FreeFile
    Open "C:\TEF_DIAL\REQ\IntPos.tmp" For Output As #lNumeroArquivo
    Print #lNumeroArquivo, "000-000 = ADM"
    Print #lNumeroArquivo, "001-000 = " & Format(gNumeroControleSolicitacao, "0000000000")
    Print #lNumeroArquivo, "999-999 = 0"
    Close #lNumeroArquivo
    If lArqTxt.FileExists("C:\TEF_DIAL\REQ\IntPos.tmp") Then
        lArqTxt.MoveFile ("C:\TEF_DIAL\REQ\IntPos.tmp"), ("C:\TEF_DIAL\REQ\IntPos.001")
    End If
    'Aguarda 7 segundos
    lHoraInicial = Time
    Do Until DateDiff("s", lHoraInicial, Time) >= 7
        If lArqTxt.FileExists("C:\TEF_DIAL\RESP\IntPos.STS") Then
            Exit Do
        End If
        DoEvents
    Loop
    
    If Not lArqTxt.FileExists("C:\TEF_DIAL\RESP\IntPos.STS") Then
        MsgBox "Gerenciador Padrão não está ativo, e será ativado automaticamente!", vbInformation, "Mensagem Padrão"
        Call AtivaGerenciadorPadrao
        'Aguarda 7 segundos
        lHoraInicial = Time
        Do Until DateDiff("s", lHoraInicial, Time) >= 7
            If lArqTxt.FileExists("C:\TEF_DIAL\RESP\IntPos.STS") Then
                Exit Do
            End If
            DoEvents
        Loop
    End If
    
    If lArqTxt.FileExists("C:\TEF_DIAL\RESP\IntPos.STS") Then
        lNumeroArquivo = FreeFile
        Open "C:\TEF_DIAL\RESP\IntPos.STS" For Input As #lNumeroArquivo
        i = 0
        Do Until EOF(lNumeroArquivo)
            Input #lNumeroArquivo, lString
            i = i + 1
            If i = 1 Then
                If lString <> "000-000 = ADM" Then
                    Exit Do
                End If
            End If
            'Verifica se o número do controle da solicitação é igual
            If Mid(lString, 1, 7) = "001-000" Then
                If Mid(lString, 11, 10) = gNumeroControleSolicitacao Then
                    lCampo001000Igual = True
                End If
            End If
        Loop
        Close #lNumeroArquivo
        lArqTxt.DeleteFile ("C:\TEF_DIAL\RESP\IntPos.STS")
        If lCampo001000Igual Then
            SolicitacaoADM = True
        Else
            MsgBox "Numero de controle de solicitação está diferente", vbInformation, "????"
        End If
    Else
        MsgBox "Gerenciador Padrão não está ativo!", vbInformation, "Mensagem Padrão"
    End If
    Exit Function
FileError:
    If Err = 76 Then
        MsgBox "Gerenciador Padrão não está instalado neste computador!", vbInformation, "Transação com TEF não aceita!"
        Exit Function
    End If
    MsgBox "ERRO não identificado!", vbInformation, "Rotina: GerenciadorPadraoAtivo"
    Exit Function
End Function
Function SolicitacaoDeCompra() As Boolean
    Dim i As Integer
    On Error GoTo FileError
    SolicitacaoDeCompra = False
    lCampo001000Igual = False
    'CRT
    'Gera Arquivo
    If lArqTxt.FileExists("C:\TEF_DIAL\REQ\IntPos.tmp") Then
        lArqTxt.DeleteFile ("C:\TEF_DIAL\REQ\IntPos.tmp")
    End If
    If lArqTxt.FileExists("C:\TEF_DIAL\REQ\IntPos.001") Then
        lArqTxt.DeleteFile ("C:\TEF_DIAL\REQ\IntPos.001")
    End If
    lNumeroArquivo = FreeFile
    Open "C:\TEF_DIAL\REQ\IntPos.tmp" For Output As #lNumeroArquivo
    Print #lNumeroArquivo, "000-000 = CRT"
    Print #lNumeroArquivo, "001-000 = " & Format(gNumeroControleSolicitacao, "0000000000")
    Print #lNumeroArquivo, "002-000 = " & Format(lNumeroCupom + 1, "000000")
    lString = Format(fValidaValor(lValorRecebido), "#########0.00")
    i = Len(lString)
    lString = Mid(lString, 1, i - 3) & Mid(lString, i - 1, 2)
    Print #lNumeroArquivo, "003-000 = " & lString
    Print #lNumeroArquivo, "999-999 = 0"
    Close #lNumeroArquivo
    If lArqTxt.FileExists("C:\TEF_DIAL\REQ\IntPos.tmp") Then
        lArqTxt.MoveFile ("C:\TEF_DIAL\REQ\IntPos.tmp"), ("C:\TEF_DIAL\REQ\IntPos.001")
    End If
    'Sleep 7000
    lHoraInicial = Time
    Do Until DateDiff("s", lHoraInicial, Time) >= 7
        If lArqTxt.FileExists("C:\TEF_DIAL\RESP\IntPos.STS") Then
            Exit Do
        End If
        DoEvents
    Loop
    If lArqTxt.FileExists("C:\TEF_DIAL\RESP\IntPos.STS") Then
        lNumeroArquivo = FreeFile
        Open "C:\TEF_DIAL\RESP\IntPos.STS" For Input As #lNumeroArquivo
        i = 0
        Do Until EOF(lNumeroArquivo)
            Input #lNumeroArquivo, lString
            i = i + 1
            If i = 1 Then
                If lString <> "000-000 = CRT" Then
                    Exit Do
                End If
            End If
            If Mid(lString, 1, 7) = "001-000" Then
                If Mid(lString, 11, 10) = gNumeroControleSolicitacao Then
                    lCampo001000Igual = True
                End If
            End If
        Loop
        Close #lNumeroArquivo
        lArqTxt.DeleteFile ("C:\TEF_DIAL\RESP\IntPos.STS")
        If lCampo001000Igual Then
            SolicitacaoDeCompra = True
        Else
            MsgBox "Numero de controle de solicitação está diferente", vbInformation, "????"
        End If
    Else
        MsgBox "TEF Não Está Ativo!", vbInformation, "Mensagem Padrão"
    End If
    Exit Function
FileError:
    MsgBox "ERRO não identificado!", vbInformation, "Rotina: SolicitacaoDeCompra"
    Exit Function
End Function
Function TestaImprimeTEF(ByVal xTipoTEF As String) As Boolean
    Dim xSaiDoLoop As Boolean
    TestaImprimeTEF = False
    xSaiDoLoop = False
    
    'Carrega Mensagem
    If Not CarregaMensagemTEF Then
        Exit Function
    End If
    
    'Imprime e Testa TEF
    Do Until xSaiDoLoop = True
        If ImprimeTEF(xTipoTEF) Then
            TestaImprimeTEF = True
            xSaiDoLoop = True
        Else
            rtxt_mensagem = "Impressora Não Responde."
            rtxt_mensagem.Visible = True
            DoEvents
            If (MsgBox("Impressora Não Responde, Tentar Imprimir Novamente ?  Sim ou Não.", vbQuestion + vbDefaultButton1 + vbYesNo, "Impressora Não Responde!") = vbNo) Then
                xSaiDoLoop = True
            Else
                rtxt_mensagem = "Aguarde!" & Chr(10) & Chr(10) & "Imprimindo TEF"
                rtxt_mensagem.Visible = True
                DoEvents
                BemaRetorno = Bematech_FI_FechaRelatorioGerencial
            End If
        End If
        xTipoTEF = "LeituraX"
        If xSaiDoLoop = True Then
            Exit Do
        End If
    Loop
End Function
Function TestaImprimeTefADM(ByVal xTipoTEF As String) As Boolean
    Dim xSaiDoLoop As Boolean
    TestaImprimeTefADM = False
    xSaiDoLoop = False
    
    'Carrega Mensagem
    If Not CarregaMensagemTEF Then
        Exit Function
    End If
    
    'Imprime e Testa TEF
    Do Until xSaiDoLoop = True
        If ImprimeTefADM(xTipoTEF) Then
            TestaImprimeTefADM = True
            xSaiDoLoop = True
        Else
            rtxt_mensagem = "Impressora Não Responde."
            rtxt_mensagem.Visible = True
            DoEvents
            If (MsgBox("Impressora Não Responde, Tentar Imprimir Novamente ?  Sim ou Não.", vbQuestion + vbDefaultButton1 + vbYesNo, "Impressora Não Responde!") = vbNo) Then
                xSaiDoLoop = True
            Else
                rtxt_mensagem = "Aguarde!" & Chr(10) & Chr(10) & "Imprimindo TEF" & Chr(10) & "Administrativo."
                rtxt_mensagem.Visible = True
                DoEvents
                BemaRetorno = Bematech_FI_FechaRelatorioGerencial
            End If
        End If
        xTipoTEF = "LeituraX"
        If xSaiDoLoop = True Then
            Exit Do
        End If
    Loop
End Function
Function TestaSolicitacaoADM() As Boolean
    Dim i As Integer
    Dim xCampo028 As Long
    On Error GoTo FileError
    xCampo028 = 0
    TestaSolicitacaoADM = False
    lCampo001000Igual = False
    rtxt_mensagem = "Aguardando Operação" & Chr(10) & "Administrativa!"
    rtxt_mensagem.Visible = True
    Do Until lArqTxt.FileExists("C:\TEF_DIAL\RESP\IntPos.001")
        DoEvents
    Loop
    rtxt_mensagem.Visible = False
    
    
    'Mostra mensagem 30
    rtxt_mensagem.Enabled = False
    If lArqTxt.FileExists("C:\TEF_DIAL\RESP\IntPos.001") Then
        lNumeroArquivo = FreeFile
        Open "C:\TEF_DIAL\RESP\IntPos.001" For Input As #lNumeroArquivo
        lMensagem = ""
        i = 0
        Do Until EOF(lNumeroArquivo)
            Input #lNumeroArquivo, lString
            If Mid(lString, 1, 7) = "001-000" Then
                If Mid(lString, 11, 10) = gNumeroControleSolicitacao Then
                    lCampo001000Igual = True
                End If
            End If
            If Mid(lString, 1, 7) = "028-000" Then
                xCampo028 = Mid(lString, 11, Len(lString) - 10)
            End If
            If Mid(lString, 1, 3) = "030" Then
                i = i + 1
                If i > 1 Then
                    lMensagem = lMensagem & Chr(10)
                End If
                lMensagem = lMensagem & Mid(lString, 11, Len(lString) - 10)
            End If
        Loop
        rtxt_mensagem = lMensagem
        rtxt_mensagem.Visible = True
        lHoraInicial = Time
        Me.Caption = "Mensagem para o Operador"
        Do Until DateDiff("s", lHoraInicial, Time) >= 5
            DoEvents
        Loop
        Me.Caption = "T.E.F."
        rtxt_mensagem.Visible = False
        Close #lNumeroArquivo
        
        If lCampo001000Igual Then
            If lMensagem Like "*OK*" Then
                If xCampo028 = 0 Then
                    lArqTxt.DeleteFile ("C:\TEF_DIAL\RESP\IntPos.001")
                Else
                    TestaSolicitacaoADM = True
                    'ImprimeTef
                End If
            Else
                'MsgBox lMensagem, vbInformation, "Mensagem de retorno"
                lArqTxt.DeleteFile ("C:\TEF_DIAL\RESP\IntPos.001")
            End If
        Else
            MsgBox "Numero de controle de solicitação está diferente", vbInformation, "????"
        End If
        'lArqTxt.DeleteFile ("C:\TEF_DIAL\RESP\IntPos.001")
    Else
        MsgBox "O arquivo IntPos.001 foi deletado de forma não identificada!", vbInformation, "Mensagem Padrão"
    End If
    Exit Function
FileError:
    MsgBox "ERRO não identificado!", vbInformation, "Rotina: SolicitacaoDeCompra"
    Exit Function
End Function


Function TestaSolicitacaoDeCompra() As Boolean
    Dim i As Integer
    Dim xCampo028 As Long
    On Error GoTo FileError
    xCampo028 = 0
    TestaSolicitacaoDeCompra = False
    lCampo001000Igual = False
    rtxt_mensagem = "Aguardando Operação!"
    rtxt_mensagem.Visible = True
    Do Until lArqTxt.FileExists("C:\TEF_DIAL\RESP\IntPos.001")
        DoEvents
    Loop
    rtxt_mensagem.Visible = False
    'lNumeroArquivo = FreeFile
    'Open "C:\TEF_DIAL\REQ\IntPos.tmp" For Output As #lNumeroArquivo
    'Print #lNumeroArquivo, "000-000 = CRT"
    'Print #lNumeroArquivo, "001-000 = " & Format(gNumeroControleSolicitacao, "0000000000")
    'Print #lNumeroArquivo, "002-000 = " & Format(lNumeroCupom + 1, "000000")
    'lString = Format(fValidaValor(lValorRecebido), "#########0.00")
    'i = Len(lString)
    'lString = Mid(lString, 1, i - 3) & Mid(lString, i - 1, 2)
    'Print #lNumeroArquivo, "003-000 = " & lString
    'Print #lNumeroArquivo, "999-999 = 0"
    'Close #lNumeroArquivo
    'If lArqTxt.FileExists("C:\TEF_DIAL\REQ\IntPos.tmp") Then
    '    lArqTxt.MoveFile ("C:\TEF_DIAL\REQ\IntPos.tmp"), ("C:\TEF_DIAL\REQ\IntPos.001")
    'End If
    'Sleep 7000
    
    
    'Mostra mensagem 30
    rtxt_mensagem.Enabled = False
    If lArqTxt.FileExists("C:\TEF_DIAL\RESP\IntPos.001") Then
        lNumeroArquivo = FreeFile
        Open "C:\TEF_DIAL\RESP\IntPos.001" For Input As #lNumeroArquivo
        lMensagem = ""
        i = 0
        Do Until EOF(lNumeroArquivo)
            Input #lNumeroArquivo, lString
            If Mid(lString, 1, 7) = "001-000" Then
                If Mid(lString, 11, 10) = gNumeroControleSolicitacao Then
                    lCampo001000Igual = True
                End If
            End If
            If Mid(lString, 1, 7) = "028-000" Then
                xCampo028 = Mid(lString, 11, Len(lString) - 10)
            End If
            If Mid(lString, 1, 3) = "030" Then
                i = i + 1
                If i > 1 Then
                    lMensagem = lMensagem & Chr(10)
                End If
                lMensagem = lMensagem & Mid(lString, 11, Len(lString) - 10)
            End If
        Loop
        rtxt_mensagem = lMensagem
        rtxt_mensagem.Visible = True
        lHoraInicial = Time
        Me.Caption = "Mensagem para o Operador"
        Do Until DateDiff("s", lHoraInicial, Time) >= 5
            DoEvents
        Loop
        Me.Caption = "T.E.F."
        rtxt_mensagem.Visible = False
        Close #lNumeroArquivo
        
        If lCampo001000Igual Then
            If lMensagem Like "*OK*" Then
                If xCampo028 = 0 Then
                    lArqTxt.DeleteFile ("C:\TEF_DIAL\RESP\IntPos.001")
                Else
                    TestaSolicitacaoDeCompra = True
                    'ImprimeTef
                End If
            Else
                'MsgBox lMensagem, vbInformation, "Mensagem de retorno"
                lArqTxt.DeleteFile ("C:\TEF_DIAL\RESP\IntPos.001")
            End If
        Else
            MsgBox "Numero de controle de solicitação está diferente", vbInformation, "????"
        End If
        'lArqTxt.DeleteFile ("C:\TEF_DIAL\RESP\IntPos.001")
    Else
        MsgBox "O arquivo IntPos.001 foi deletado de forma não identificada!", vbInformation, "Mensagem Padrão"
    End If
    Exit Function
FileError:
    MsgBox "ERRO não identificado!", vbInformation, "Rotina: SolicitacaoDeCompra"
    Exit Function
End Function


