VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frm_preview 
   Caption         =   "Preview de Impressão"
   ClientHeight    =   6105
   ClientLeft      =   2265
   ClientTop       =   6960
   ClientWidth     =   8085
   Icon            =   "preview.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6105
   ScaleWidth      =   8085
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmd_calc 
      Height          =   670
      Left            =   120
      Picture         =   "preview.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Calculadora."
      Top             =   5340
      Width           =   600
   End
   Begin VB.CommandButton cmd_marcar 
      Default         =   -1  'True
      Height          =   670
      Left            =   780
      Picture         =   "preview.frx":15E4
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Destaca texto selecionado."
      Top             =   5340
      Width           =   600
   End
   Begin VB.CommandButton cmd_localizar 
      Height          =   670
      Left            =   1440
      Picture         =   "preview.frx":1A26
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Pesquisa texto do relatório."
      Top             =   5340
      Width           =   600
   End
   Begin VB.CommandButton cmd_enviar_email 
      Height          =   670
      Left            =   2760
      Picture         =   "preview.frx":2E98
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Envia este relatório por E-Mail."
      Top             =   5340
      Width           =   600
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Height          =   670
      Left            =   3420
      Picture         =   "preview.frx":4172
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   5340
      Width           =   600
   End
   Begin VB.CommandButton cmd_imprimir 
      Height          =   670
      Left            =   2100
      Picture         =   "preview.frx":5804
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Imprime este relatório."
      Top             =   5340
      Width           =   600
   End
   Begin RichTextLib.RichTextBox txt_preview 
      Height          =   5235
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   9234
      _Version        =   393217
      BackColor       =   16777215
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      RightMargin     =   9,99999e5
      TextRTF         =   $"preview.frx":6E0E
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
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   5040
      Top             =   5460
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   4380
      Top             =   5460
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
End
Attribute VB_Name = "frm_preview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dados As String
Dim tbl_relatorio As Table
Dim lCurrentY As Currency
Dim l_status As Integer
Dim lNomeArquivo As String
Dim lNomeArquivoHTML As String
Dim lString As String
Dim lNomeConta As String
Dim lSenha As String
Dim lEMailRemetente As String
Dim lEMailDestinatario As String
Dim lTituloRelatorio As String
Dim lResultado As Long
Dim lPesquisa As String

Dim AcionadoEnviarEmail As Boolean
Private Declare Function InternetAutodial Lib "wininet.dll" (ByVal dwflags As Long, ByVal dwReserved As Long) As Long
Private Declare Function InternetAutodialHangup Lib "wininet.dll" (ByVal dwReserved As Long) As Long
Private Const INTERNET_AUTODIAL_FORCE_ONLINE = 1
Private Const INTERNET_AUTODIAL_FORCE_UNATTENDED = 2

'Funçoes Utilizadas Para desabilitar o menu(ControlBox) do windows
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Const MF_BYPOSITION = &H400&
Private Sub cmd_calc_Click()
    Dim retval As Long
    retval = Shell("calc", vbNormalFocus)
End Sub
Private Sub cmd_enviar_email_Click()
    Call InformaDestinatario
    If CriaArquivoHTML Then
        Call GravaAuditoria(1, Me.name, 8, "")
        Call EnviaEmail
    End If
    'Indicando ao usuário a conexão
    'If InternetAutodial(INTERNET_AUTODIAL_FORCE_ONLINE, 0) Then
    '    MsgBox "Você esta conectado!", vbInformation
    'inicia a discagem automaticamente
    'ElseIf InternetAutodial(INTERNET_AUTODIAL_FORCE_UNATTENDED, 0) Then
    '    MsgBox "Você já esta conectado!", vbInformation
    'End If
    ChamaDrive
End Sub
Private Sub cmd_imprimir_Click()
    Mid(g_string, 1, 1) = "1"
    Mid(g_string, 13, 1) = "P"
    Call GravaAuditoria(1, Me.name, 7, "")
    'frm_seleciona_impressora.Show 1
    Call Relatorio
    cmd_sair.SetFocus
End Sub

Private Sub cmd_localizar_Click()
    Dim xVariavel As String
    Dim xNumero As Integer
    
    xVariavel = g_string
    g_string = lPesquisa
    localizar.Show 1
    If g_string <> "" Then
        lPesquisa = RetiraGString(1)
        If RetiraGString(2) = "True" Then
            xNumero = 4
        Else
            xNumero = 0
        End If
        If lResultado = txt_preview.SelStart Then
            txt_preview.SelStart = txt_preview.SelStart + 1
        End If
        lResultado = txt_preview.Find(lPesquisa, txt_preview.SelStart, Len(txt_preview.Text), xNumero)
        If lResultado = -1 Then
            MsgBox "Não é possível localizar " & Chr(39) & lPesquisa & Chr(39), vbExclamation, "Pesquisa de texto."
        Else
            'g_string = xVariavel
            txt_preview.SetFocus
            'Exit Sub
        End If
    End If
    g_string = xVariavel
    Exit Sub
End Sub
Private Sub cmd_marcar_Click()
    If txt_preview.SelLength > 0 Then
        If txt_preview.SelBold = False Then
            txt_preview.SelBold = True
            txt_preview.SelColor = 16711680
        Else
            txt_preview.SelBold = False
            txt_preview.SelColor = 0
        End If
    End If
End Sub
Private Sub cmd_sair_Click()
    HabilitaMenu
    Unload Me
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    If l_status = 0 Then
        If cmd_sair.Visible = True Then
            cmd_sair.SetFocus
        End If
        l_status = 1
        If Mid(g_string, 1, 1) = "1" Then
            Mid(g_string, 13, 1) = "P"
            Relatorio
            cmd_sair_Click
        End If
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then
        KeyCode = 0
        cmd_localizar_Click
    ElseIf KeyCode = vbKeyF10 Then
        KeyCode = 0
        cmd_imprimir_Click
    End If
End Sub

Private Sub Form_Load()
    CentraForm Me
    Screen.MousePointer = 1
    l_status = 0
    AcionadoEnviarEmail = False
    lTituloRelatorio = RetiraGString(2)
    g_string = RetiraGString(1)
    If Mid(g_string, Len(g_string) - 3, 4) = ".LOG" Then
        AtualizaPagina
    Else
        If Mid(g_string, 1, 1) = "0" Then
            Mid(g_string, 13, 1) = "V"
            AtualizaPagina
        Else
           '// desabilita o botão fechar
            DesabilitaMenu
'            cmd_calc.Visible = False
'            cmd_enviar_email.Visible = False
'            cmd_imprimir.Visible = False
'            cmd_localizar.Visible = False
'            cmd_marcar.Visible = False
'            cmd_sair.Visible = False
            
            'Me.ControlBox = False
        End If
    End If
End Sub

Private Sub DesabilitaMenu()
  Dim hMenu As Long
  hMenu = GetSystemMenu(hWnd, False)
  DeleteMenu hMenu, 6, MF_BYPOSITION

  cmd_calc.Visible = False
  cmd_enviar_email.Visible = False
  cmd_imprimir.Visible = False
  cmd_localizar.Visible = False
  cmd_marcar.Visible = False
  cmd_sair.Visible = False
End Sub
Private Sub HabilitaMenu()
  Dim hMenu As Long
  hMenu = GetSystemMenu(hWnd, True)
  DeleteMenu hMenu, 6, MF_BYPOSITION

  cmd_calc.Visible = True
  cmd_enviar_email.Visible = True
  cmd_imprimir.Visible = True
  cmd_localizar.Visible = True
  cmd_marcar.Visible = True
  cmd_sair.Visible = True
End Sub

Private Sub AtualizaPagina()
    Dim i As Integer
    lNomeArquivo = Mid(g_string, 2, Len(g_string) - 1)
    If Mid(lNomeArquivo, Len(lNomeArquivo) - 3, 4) = ".LOG" Then
        'Open Mid(g_string, 2, Len(g_string) - 1) For Input As #1
        'txt_preview.Text = StrConv(InputB$(LOF(1), 1), vbUnicode)
        Set gArquivoTMP = gArqTxt.OpenTextFile(lNomeArquivo, ForReading)
        txt_preview.Text = gArquivoTMP.ReadAll
    Else
        'Open "D:\MEUS DOCUMENTOS\MARCOS\RESC-MAR.TXT" For Input As #1
        'Do Until EOF(1)
        '    DoEvents
        '    Line Input #1, dados
        '    'txt_preview.SelColor = &H0&
        '    'txt_dados.SelColor = &HFF&
        '    If Mid(dados, 1, 2) <> "@@" Then
        '        txt_preview.Text = txt_preview.Text + dados + Chr(13) + Chr(10)
        '    ElseIf dados = "@@Printer.NewPage" Then
        '        txt_preview.Text = txt_preview.Text + " " + Chr(13) + Chr(10)
        '    ElseIf dados = "@@Printer.EndDoc" Then
        '        Exit Do
        '    End If
        'Loop
        
        
        
        txt_preview.Visible = False
        'Open Mid(g_string, 2, Len(g_string) - 1) For Input As #1
        'txt_preview.Text = StrConv(InputB$(LOF(1), 1), vbUnicode)
        Set gArquivoTMP = gArqTxt.OpenTextFile(lNomeArquivo, ForReading)
        txt_preview.Text = gArquivoTMP.ReadAll
        'txt_preview.TextRTF = StrConv(InputB$(LOF(1), 1), vbUnicode)
        'Do Until EOF(1)
        '    DoEvents
        '    Line Input #1, dados
        '    'txt_preview.SelColor = &H0&
        '    'txt_dados.SelColor = &HFF&
        '    If Mid(dados, 1, 2) <> "@@" Then
        '        txt_preview.Text = txt_preview.Text + Mid(dados, 16, Len(dados) - 15) + Chr(13) + Chr(10)
        '    ElseIf dados = "@@Printer.NewPage" Then
        '        txt_preview.Text = txt_preview.Text + " " + Chr(13) + Chr(10)
        '    ElseIf dados = "@@Printer.EndDoc" Then
        '        Exit Do
        '    End If
        'Loop
    End If
    'Close #1
    gArquivoTMP.Close
    txt_preview.Visible = True
End Sub
Function CriaArquivoHTML() As Boolean
    CriaArquivoHTML = False
    'Atribui nome do arquivo TMV para variável lNomeArquivo
    lNomeArquivo = Mid(g_string, 2, Len(g_string) - 1)
    
    'Atribui nome do arquivo HTML para variável lNomeArquivo
    'através do arquivo TMV, substituindo a extensão
    lNomeArquivoHTML = Mid(lNomeArquivo, 1, Len(lNomeArquivo) - 3) & "HTML"
    
    'Verifica se o Arquivo TMV existe
    If gArqTxt.FileExists(lNomeArquivo) Then
        
        'Abre o Arquivo TMV para leitura
        Set gArquivoTMV = gArqTxt.OpenTextFile(lNomeArquivo, ForReading)
        
        'Cria o Arquivo HTML
        Set gArquivoHTML = gArqTxt.CreateTextFile(lNomeArquivoHTML, True)
        
        'Inicia formato HTML
        lString = "<HTML><HEAD><TITLE>"
        gArquivoHTML.WriteLine (lString)
        
        'Nomeia título do HTML
        lString = lTituloRelatorio
        gArquivoHTML.WriteLine (lString)
        
        'Formata o início do corpo do HTML
        lString = "</TITLE></HEAD><BODY BGCOLOR=#C0D9D9><FONT face=Courier New size=3><PRE>"
        gArquivoHTML.WriteLine (lString)
        
        'Loop de leitura do arquivo TMV
        Do Until gArquivoTMV.AtEndOfStream
            
            'Lê linha do arquivo TMV
            lString = gArquivoTMV.ReadLine
            
            'Grava linha do arquivo HTML
            gArquivoHTML.WriteLine (lString)
        Loop
        
        'Finaliza o corpo do HTML
        lString = "</PRE></FONT></BODY></HTML>"
        gArquivoHTML.WriteLine (lString)
    Else
        MsgBox "Arquivo Inexistente", vbInformation, "Erro de Integridade"
        Exit Function
    End If
    'Fecha arquivos TMV e HTML
    gArquivoTMV.Close
    gArquivoHTML.Close
    CriaArquivoHTML = True
End Function
Function EnviaEmail() As Boolean
    EnviaEmail = False
    
    MAPISession1.UserName = lNomeConta '"tasso_cerrado@uol.com.br"
    MAPISession1.Password = lSenha '"lara28"
    MAPISession1.SignOn
    
    
    With MAPIMessages1
        .SessionID = MAPISession1.SessionID
        .Compose
        .RecipAddress = lEMailDestinatario '"tassoteixeira@uol.com.br"
        .MsgSubject = "SGP Cerrado - " & g_nome_empresa & ": " & lTituloRelatorio
        .MsgNoteText = "Segue em anexo o " & lTituloRelatorio & Chr(10) & "Da Empresa: " & g_nome_empresa
        
        'anexa no final da mensagem
        .AttachmentPosition = Len(MAPIMessages1.MsgNoteText)
        
        'define o tipo de dados do anexo
        .AttachmentType = mapData
        
        'da um nome ao anexo
        .AttachmentName = lNomeArquivoHTML
        
        'define o caminho e nome do arquivo a anexar
        .AttachmentPathName = gDrive & gDiretorioData & lNomeArquivoHTML
    
        'envia o arquivo
        .send
    End With
    MAPISession1.SignOff
    
    gArqTxt.DeleteFile (gDrive & gDiretorioData & lNomeArquivoHTML)
    
    EnviaEmail = True
End Function
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
End Sub
Function InformaDestinatario() As Boolean
    InformaDestinatario = False
    
    lNomeArquivo = "Relacao_Email.txt"
    If Not gArqTxt.FileExists(lNomeArquivo) Then
        Set gArquivoHTML = gArqTxt.CreateTextFile(lNomeArquivo, True)
        lString = InputBox("Informe o nome da conta: ", "Nome da Conta")
        gArquivoHTML.WriteLine ("Nome da Conta: = " & lString)
        lString = InputBox("Informe a senha: ", "Senha")
        gArquivoHTML.WriteLine ("Senha: = " & lString)
        lString = InputBox("Informe o E-Mail do remetente: ", "E-Mail do Remetente")
        gArquivoHTML.WriteLine ("E-Mail Remetente: = " & lString)
        lString = InputBox("Informe o E-Mail do Destinatário (01): ", "E-Mail do Destinatário")
        gArquivoHTML.WriteLine ("E-Mail Destinatario (01): = " & lString)
        gArquivoHTML.Close
    End If
    
    Set gArquivoHTML = gArqTxt.OpenTextFile(lNomeArquivo, ForReading)
        
        
    Do Until gArquivoHTML.AtEndOfStream
        lString = gArquivoHTML.ReadLine
        If lString Like "*Nome da Conta*" Then
            lNomeConta = Mid(lString, 18, Len(lString) - 17)
        ElseIf lString Like "*Senha*" Then
            lSenha = Mid(lString, 10, Len(lString) - 9)
        ElseIf lString Like "*E-Mail Remetente*" Then
            lEMailRemetente = Mid(lString, 21, Len(lString) - 20)
        ElseIf lString Like "*E-Mail Destinatario*" Then
            lEMailDestinatario = Mid(lString, 29, Len(lString) - 28)
        End If
    Loop
    gArquivoHTML.Close
    InformaDestinatario = True
End Function
Private Sub RptDefinePaperSize(ByVal pTamanho As String)
    On Error GoTo FileError
        
    Printer.PaperSize = pTamanho
    
    Exit Sub
FileError:
    'MsgBox "Erro ao definir tamanho do Papel da Impressora!", vbInformation, "Erro de Configuração"
    Exit Sub
End Sub
Private Sub Relatorio()
    'Open Mid(g_string, 2, Len(g_string) - 1) For Input As #1
    lNomeArquivo = Mid(g_string, 2, Len(g_string) - 1)
    Set gArquivoTMP = gArqTxt.OpenTextFile(lNomeArquivo, ForReading)
    'Do Until EOF(1)
    Do Until gArquivoTMP.AtEndOfStream
        DoEvents
        'Line Input #1, dados
        dados = gArquivoTMP.ReadLine
        'txt_preview.SelColor = &H0&
'Print #1, Chr(27) & "0" ' Muda o passo p/ 8 LPP
        If Mid(dados, 1, 19) = "@@Printer.ScaleMode" Then
            Printer.ScaleMode = Mid(dados, 23, Len(dados) - 22)
        ElseIf Mid(dados, 1, 19) = "@@Printer.PaperSize" Then
            RptDefinePaperSize (Mid(dados, 23, Len(dados) - 22))
        ElseIf Mid(dados, 1, 18) = "@@Printer.FontName" Then
            If g_impressora_matricial = True Then
                If UCase(g_nome_empresa) Like "*POUSO ALTO*" Then
                    If Mid(dados, 22, Len(dados) - 21) = "Draft 10cpi" Then
                        Printer.FontName = "Sans Serif 10cpi"
                        Printer.FontBold = False
                    Else
                        Printer.FontName = Mid(dados, 22, Len(dados) - 21)
                    End If
                'ElseIf UCase(g_nome_empresa) Like "*LUDOVICO*" Or UCase(g_nome_empresa) Like "*LUDOVICO*" Then
                '    If Mid(dados, 22, Len(dados) - 21) = "Draft 10cpi" Then
                '        Printer.FontName = "Courier"
                '        Printer.FontSize = 10
                '        Printer.FontBold = False
                '    Else
                '        Printer.FontName = Mid(dados, 22, Len(dados) - 21)
                '    End If
                Else
            
            
                    If Mid(dados, 22, Len(dados) - 21) = "Sans Serif 17cpi" And g_tamanho_impressora = 80 Then
                        Printer.FontName = Mid(dados, 22, Len(dados) - 21)
                        Printer.FontBold = False
                    Else
                        Printer.FontName = Mid(dados, 22, Len(dados) - 21)
                    End If
                
                End If
            
            Else
                If Mid(dados, 22, Len(dados) - 21) = "Sans Serif 17cpi" Then
                    Printer.FontName = "Courier New"
                    Printer.FontSize = 6.55
                ElseIf Mid(dados, 22, Len(dados) - 21) = "Courier New" Then
                    Printer.FontName = Mid(dados, 22, Len(dados) - 21)
                ElseIf Mid(dados, 22, Len(dados) - 21) = "Draft 10cpi" Then
                    Printer.FontName = "Courier New"
                    Printer.FontSize = 11
                ElseIf Mid(dados, 22, Len(dados) - 21) = "Lucida Console 8cpi" Then
                    Printer.FontName = "Courier New"
                    Printer.FontSize = 8
                ElseIf Mid(dados, 22, Len(dados) - 21) = "Lucida Console 7cpi" Then
                    Printer.FontName = "Courier New"
                    Printer.FontSize = 7
                ElseIf Mid(dados, 22, Len(dados) - 21) = "Lucida Console 6cpi" Then
                    Printer.FontName = "Courier New"
                    Printer.FontSize = 6
                Else
                    Printer.FontName = "Courier New"
                    Printer.FontSize = 10
                End If
            End If
        ElseIf Mid(dados, 1, 18) = "@@Printer.FontSize" Then
            Printer.FontSize = Mid(dados, 22, Len(dados) - 21)
        ElseIf Mid(dados, 1, 18) = "@@Printer.FontBold" Then
            If Printer.FontName = "Sans Serif 17cpi" And g_tamanho_impressora = 80 Then
                Printer.FontBold = False
            Else
                Printer.FontBold = Mid(dados, 22, Len(dados) - 21)
            End If
        ElseIf Mid(dados, 1, 14) = "@@ImprimeTexto" Then
            Call RelatorioImprimeTexto(dados)
        ElseIf Mid(dados, 1, 18) = "@@Printer.CurrentY" Then
            If Mid(dados, 22, Len(dados) - 21) = "y_local" Then
                Printer.CurrentY = lCurrentY
            Else
                Printer.CurrentY = Mid(dados, 22, Len(dados) - 21)
            End If
        ElseIf Mid(dados, 1, 15) = "@@Printer.Print" Then
            Printer.Print Mid(dados, 18, Len(dados) - 18)
        ElseIf Mid(dados, 1, 14) = "@Printer.Print" Then
            Printer.Print Mid(dados, 16, Len(dados) - 15)
        ElseIf Mid(dados, 1, 9) = "@@y_local" Then
            Call RelatorioBuscaCoordenadaY
        ElseIf Mid(dados, 1, 17) = "@@Printer.NewPage" Then
            Printer.NewPage
        ElseIf Mid(dados, 1, 16) = "@@Printer.EndDoc" Then
            Printer.EndDoc
            Exit Do
        Else
            MsgBox "Comando não interpretado! " & Chr(10) & dados
        End If
    Loop
    'Close #1
    gArquivoTMP.Close
End Sub
Private Sub RelatorioBuscaCoordenadaY()
    lCurrentY = Printer.CurrentY
End Sub
Private Sub RelatorioImprimeTexto(x_dados As String)
    Dim i As Integer
    Dim i2 As Integer
    Dim x_numero1 As Currency
    Dim x_numero2 As Currency
    Dim x_numero3 As Currency
    Dim x_numero4 As Currency
    Dim x_string As String
    For i = 17 To Len(x_dados)
        If Mid(x_dados, i, 1) = Chr(34) Then
            x_string = Mid(x_dados, 17, i - 17)
            Exit For
        End If
    Next
    i = i + 3
    For i2 = i To Len(x_dados)
        If Mid(x_dados, i2, 1) = "," Then
            x_numero1 = Val(Mid(x_dados, i, i2 - i))
            Exit For
        End If
    Next
    i = i2 + 2
    For i2 = i To Len(x_dados)
        If Mid(x_dados, i2, 1) = "," Then
            x_numero2 = Val(Mid(x_dados, i, i2 - i))
            Exit For
        End If
    Next
    i = i2 + 2
    For i2 = i To Len(x_dados)
        If Mid(x_dados, i2, 1) = "," Then
            x_numero3 = Val(Mid(x_dados, i, i2 - i))
            Exit For
        End If
    Next
    i = i2 + 2
    x_numero4 = Val(Mid(x_dados, i, Len(x_dados) + 1 - i))
    ImprimeTexto x_string, x_numero1, x_numero2, x_numero3, x_numero4
End Sub
Private Sub Form_Resize()
    txt_preview.Width = Me.Width - 165
    txt_preview.Height = Me.Height - 1030
    
    cmd_calc.Top = Me.Height - 1070
    cmd_marcar.Top = Me.Height - 1070
    cmd_localizar.Top = Me.Height - 1070
    cmd_imprimir.Top = Me.Height - 1070
    cmd_enviar_email.Top = Me.Height - 1070
    cmd_sair.Top = Me.Height - 1070

    cmd_calc.Left = Me.Width / 2 - 1750
    cmd_marcar.Left = Me.Width / 2 - 1050
    cmd_localizar.Left = Me.Width / 2 - 350
    cmd_imprimir.Left = Me.Width / 2 + 350
    cmd_enviar_email.Left = Me.Width / 2 + 1050
    cmd_sair.Left = Me.Width / 2 + 1750
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
    If Len(g_string) > 0 Then
        If Mid(g_string, Len(g_string) - 3, 4) <> ".LOG" Then
            Mid(g_string, 13, 1) = "V"
            
            If gArqTxt.FileExists(gDrive & gDiretorioData & Mid(g_string, 2, Len(g_string) - 1)) Then
                gArqTxt.DeleteFile (gDrive & gDiretorioData & Mid(g_string, 2, Len(g_string) - 1))
            End If
            
            Mid(g_string, 13, 1) = "P"
            
            If gArqTxt.FileExists(gDrive & gDiretorioData & Mid(g_string, 2, Len(g_string) - 1)) Then
                gArqTxt.DeleteFile (gDrive & gDiretorioData & Mid(g_string, 2, Len(g_string) - 1))
            End If
            
        End If
    End If
    If AcionadoEnviarEmail Then
        If InternetAutodialHangup(0) Then
            MsgBox "Você esta desconectado", vbInformation
        End If
    End If
    g_string = ""
End Sub

