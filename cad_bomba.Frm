VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form cadastro_bomba 
   Caption         =   "Cadastro de Bicos de Bombas de Combustíveis"
   ClientHeight    =   4305
   ClientLeft      =   75
   ClientTop       =   4020
   ClientWidth     =   6990
   Icon            =   "cad_bomba.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   Picture         =   "cad_bomba.frx":030A
   ScaleHeight     =   4305
   ScaleWidth      =   6990
   Begin MSCommLib.MSComm MSCommAutomacao 
      Left            =   1020
      Top             =   3660
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Timer TimerAutomacao 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   120
      Top             =   3780
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   2820
      Picture         =   "cad_bomba.frx":0750
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   3360
      Width           =   795
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   1920
      Picture         =   "cad_bomba.frx":1DE2
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Exclui o registro atual."
      Top             =   3360
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   1020
      Picture         =   "cad_bomba.frx":3474
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Altera o registro atual."
      Top             =   3360
      Width           =   795
   End
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   120
      Picture         =   "cad_bomba.frx":496E
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Cria um novo registro."
      Top             =   3360
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   3195
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   6735
      Begin VB.TextBox txtNumeroSerie 
         Height          =   300
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   6
         Top             =   1080
         Width           =   1335
      End
      Begin MSAdodcLib.Adodc adodc_produto 
         Height          =   330
         Left            =   2700
         Top             =   2820
         Visible         =   0   'False
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "adodc_produto"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc adodc_combustivel 
         Height          =   330
         Left            =   2700
         Top             =   660
         Visible         =   0   'False
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "adodc_combustivel"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataListLib.DataCombo dtcbo_combustivel 
         Bindings        =   "cad_bomba.frx":6000
         Height          =   315
         Left            =   1680
         TabIndex        =   4
         Top             =   660
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   "dtcbo_combustivel"
      End
      Begin VB.ComboBox cbo_tipo_preco 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   2340
         Width           =   1455
      End
      Begin VB.TextBox txt_numero_ilha 
         Height          =   300
         Left            =   5280
         MaxLength       =   1
         TabIndex        =   14
         Top             =   1920
         Width           =   255
      End
      Begin VB.TextBox txt_tanque 
         Height          =   300
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   12
         Top             =   1920
         Width           =   255
      End
      Begin VB.TextBox msk_venda 
         Height          =   300
         Left            =   5280
         MaxLength       =   10
         TabIndex        =   10
         Top             =   1500
         Width           =   1335
      End
      Begin VB.TextBox msk_custo 
         Height          =   300
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   8
         Top             =   1500
         Width           =   1335
      End
      Begin VB.TextBox txt_codigo 
         Height          =   300
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
      Begin MSDataListLib.DataCombo dtcbo_produto 
         Bindings        =   "cad_bomba.frx":6020
         Height          =   315
         Left            =   1680
         TabIndex        =   18
         Top             =   2760
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   "dtcbo_produto"
      End
      Begin VB.Label Label8 
         Caption         =   "&Número de Série"
         Height          =   300
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Produto"
         Height          =   300
         Left            =   120
         TabIndex        =   17
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo de Preço"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   2340
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Número da &Ilha"
         Height          =   300
         Left            =   3720
         TabIndex        =   13
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Número do &Tanque"
         Height          =   300
         Left            =   120
         TabIndex        =   11
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "&Número do Bico"
         Height          =   300
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Co&mbustível"
         Height          =   300
         Left            =   120
         TabIndex        =   3
         Top             =   660
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "&Preço de Custo"
         Height          =   300
         Left            =   120
         TabIndex        =   7
         Top             =   1500
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Preço de &Venda"
         Height          =   300
         Left            =   3720
         TabIndex        =   9
         Top             =   1500
         Width           =   1455
      End
   End
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   4680
      TabIndex        =   25
      Top             =   3240
      Width           =   2175
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "cad_bomba.frx":603C
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "cad_bomba.frx":7536
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "cad_bomba.frx":8A30
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "cad_bomba.frx":9EA2
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   6060
      Picture         =   "cad_bomba.frx":B424
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Cancela o registro atual."
      Top             =   3360
      Width           =   795
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   5160
      Picture         =   "cad_bomba.frx":C91E
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Confirma o registro atual."
      Top             =   3360
      Width           =   795
   End
End
Attribute VB_Name = "cadastro_bomba"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lFlagCadastro
Dim lOpcao As Integer
Dim lCodigo As Integer
Dim lPrecoVenda As Currency

Dim lAutomacaoPorta As Integer
Dim lAutomacaoVelocidade As String
Dim lAutomacaoDtr As Boolean
Dim lAutomacaoRts As Boolean
Dim lAutomacaoFlag As Integer
Dim lAutomacaoBicoEncerrante As Integer
Dim lAutomacaoAbertura As Currency
Dim lAutomacaoEncerrante As Currency
Dim lAutomacaoData As Date
Dim lAutomacaoPeriodo As Integer
Dim lAutomacaoIlha As Integer

Private BicoAutomacao As New cBicoAutomacao
Private Bomba As New cBomba

Dim MyAppID
Private Sub AtivaBotoes()
    cmd_novo.Enabled = True
    cmd_alterar.Enabled = True
    cmd_excluir.Enabled = True
    cmd_sair.Enabled = True
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
    frm_move.Visible = True
    txt_codigo.Enabled = True
End Sub
Private Sub AtualTabe()
    If lOpcao = 1 Then
        Bomba.Empresa = g_empresa
        Bomba.Codigo = Val(txt_codigo.Text)
    End If
    Bomba.TipoCombustivel = dtcbo_combustivel.BoundText
    Bomba.NumeroSerie = txtNumeroSerie.Text
    Bomba.PrecoCusto = fValidaValor(msk_custo.Text)
    Bomba.PrecoVenda = fValidaValor(msk_venda.Text)
    Bomba.NumeroTanque = txt_tanque.Text
    Bomba.NumeroIlha = txt_numero_ilha.Text
    If cbo_tipo_preco.ListIndex = 0 Then
        Bomba.TipoPreco = "V"
    ElseIf cbo_tipo_preco.ListIndex = 1 Then
        Bomba.TipoPreco = "P"
    End If
    Bomba.CodigoProduto = dtcbo_produto.BoundText
End Sub
Private Sub AtualTela()
    lCodigo = Bomba.Codigo
    lPrecoVenda = Bomba.PrecoVenda
    txt_codigo.Text = Bomba.Codigo
    dtcbo_combustivel.BoundText = ""
    dtcbo_combustivel.BoundText = Bomba.TipoCombustivel
    txtNumeroSerie.Text = Bomba.NumeroSerie
    msk_custo.Text = Format(Bomba.PrecoCusto, "###0.0000")
    msk_venda.Text = Format(Bomba.PrecoVenda, "###0.0000")
    txt_tanque.Text = Bomba.NumeroTanque
    txt_numero_ilha.Text = Bomba.NumeroIlha
    cbo_tipo_preco.ListIndex = -1
    If Bomba.TipoPreco = "V" Then
        cbo_tipo_preco.ListIndex = 0
    ElseIf Bomba.TipoPreco = "P" Then
        cbo_tipo_preco.ListIndex = 1
    End If
    dtcbo_produto.BoundText = ""
    dtcbo_produto.BoundText = Bomba.CodigoProduto
    frm_dados.Enabled = False
End Sub
Public Sub AutomacaoAbrePorta()
    If MSCommAutomacao.PortOpen = False Then
        MSCommAutomacao.CommPort = lAutomacaoPorta
        MSCommAutomacao.Settings = lAutomacaoVelocidade
        MSCommAutomacao.PortOpen = True
    End If
End Sub
Function ComunicaAutomacaoCerrado(ByVal pComando As String, ByVal pParametro As String) As Boolean
    Dim xArquivoTmp As String
    Dim xArquivoPedido As String
    Dim xArquivoResp As String
    Dim xComputadorAutomacao As String
    Dim xHoraInicial As Date
    Dim xComando As String
    Dim xRetorno As String
    Dim xParametro As String
    Dim xTempo As Integer
    Dim xFaseErro As Integer

    On Error GoTo FileError

    ComunicaAutomacaoCerrado = False
    
    'Pega o NOME do computador que tem Ligado Fisicamente o Equipamento de Automação
    xFaseErro = 1
    xComputadorAutomacao = ReadINI("LOCALIZACAO", "Computador com automacao", "C:\Cerrado\AutoCerrado\AutoCerrado.INI")
    If xComputadorAutomacao = "" Then
        xComputadorAutomacao = GetIPHostName()
    End If
    
    'Monta nome do Computador + Diretório + Arquivo do Pedido de comunicação
    'Ex: \\Servidor\Automacao\Pedido_ddmmyyyy_HHmmss.TMP
    xFaseErro = 10
    xArquivoTmp = "\\" & xComputadorAutomacao & "\Automacao\Pedido_" & Format(Date, "ddmmyyyy") & "_" & Format(Time, "HHmmss") & ".TMP"
    xArquivoPedido = Mid(xArquivoTmp, 1, Len(xArquivoTmp) - 3) & "AUT"

    'Cria o arquivo .TMP de comunicação
    xFaseErro = 20
    Set gArquivoTXT = gArqTxt.CreateTextFile(xArquivoTmp)
    gArquivoTXT.WriteLine ("[PEDIDO AUTOMACAO]")
    gArquivoTXT.WriteLine ("Comando=" & pComando)
    gArquivoTXT.WriteLine ("Origem=" & GetIPHostName())
    gArquivoTXT.WriteLine ("Parametro=" & pParametro)
    gArquivoTXT.Close

    'Renomeia arquivo .TMP para .AUT
    xFaseErro = 30
    If gArqTxt.FileExists(xArquivoTmp) Then
        gArqTxt.MoveFile (xArquivoTmp), (xArquivoPedido)
    End If
    
    'Monta nome do Arquivo de Retorno
    xArquivoResp = "C:\Automacao\Retorno_" & Mid(xArquivoPedido, Len(xArquivoPedido) - 19 + 1, 19)

    xFaseErro = 40
    xTempo = 7
    'Aguarda até 7 Segundos para o retorno
    xHoraInicial = Time
    Do Until DateDiff("s", xHoraInicial, Time) >= xTempo
        If gArqTxt.FileExists(xArquivoResp) Then
            Exit Do
        End If
        DoEvents
    Loop
    
    'Verifica se o Retorno existe
    xFaseErro = 50
    If gArqTxt.FileExists(xArquivoResp) Then
        xFaseErro = 51
        'Existindo lê o retorno
        xComando = ReadINI("RETORNO AUTOMACAO", "Comando", xArquivoResp)
        xFaseErro = 52
        xRetorno = ReadINI("RETORNO AUTOMACAO", "Retorno", xArquivoResp)
        xFaseErro = 53
        xParametro = ReadINI("RETORNO AUTOMACAO", "Parametro", xArquivoResp)
        'Deleta o arquivo de retorno retorno
        xFaseErro = 54
        gArqTxt.DeleteFile (xArquivoResp)
        xFaseErro = 55
        If xRetorno = "OK" Then
            xFaseErro = 56
            ComunicaAutomacaoCerrado = True
        Else
            xFaseErro = 57
        End If
    Else
        xFaseErro = 70
        'MsgBox "arquivo nao encontrado=" & xArquivoResp
        'Deleta o Arquivo de pedido
        'Pois fica sub-entendido que o mesmo ainda existe
        gArqTxt.DeleteFile (xArquivoPedido)
        xFaseErro = 71
    End If
    xFaseErro = 100
    
    Exit Function

FileError:
    Call CriaLogCupom(Time & " - Erro ComunicaAutomacaoCerrado: xFaseErro=" & xFaseErro & " Erro=" & Err.Number & " - " & Err.Description)
    MsgBox "Erro ao tentar comunicação com o programa AutoCerrado!" & vbCrLf & "xFaseErro=" & xFaseErro & vbCrLf & "xArquivoResp=" & xArquivoResp & vbCrLf & "xArquivoPedido=" & xArquivoPedido, vbCritical, "Erro na Automação de Bomba!"
    Exit Function
End Function
Function ConverteBico(ByVal pBico As String) As String
    ConverteBico = pBico
    If BicoAutomacao.LocalizarCodigoBico(Val(pBico)) Then
        ConverteBico = BicoAutomacao.CodigoAutomacao
        Call CriaLogCupom(Time & " - ConverteBico: Bico a converter=" & pBico & " - Bico Convertido:" & BicoAutomacao.CodigoAutomacao)
    Else
        Call CriaLogCupom(Time & " - Erro ConverteBico: Bico nao localizado=" & pBico)
    End If
'    If UCase(g_nome_empresa) Like "*BOSQUE*" Then
'        If pBico = "01" Then
'            ConverteBico = "08"
'        ElseIf pBico = "02" Then
'            ConverteBico = "09"
'        ElseIf pBico = "03" Then
'            ConverteBico = "04"
'        ElseIf pBico = "04" Then
'            ConverteBico = "05"
'        ElseIf pBico = "05" Then
'            ConverteBico = "0C"
'        ElseIf pBico = "06" Then
'            ConverteBico = "0D"
'        ElseIf pBico = "07" Then
'            ConverteBico = "06"
'        ElseIf pBico = "08" Then
'            ConverteBico = "07"
'        End If
'    End If
End Function



Public Sub AutomacaoBuscaPreco(ByVal pBico As String)
    Dim i As Integer
    Dim x_string As String
    Dim x_quant As Integer
    'Le Preço Atual
    x_quant = MSCommAutomacao.InBufferCount
    Call CriaLogCupom(Time & " - Leitura de Preço pela Automação: x_quant=" & x_quant)
    If x_quant = 5 Then
        x_string = MSCommAutomacao.Input
        Call CriaLogCupom(Time & " - Erro na Leitura de Preço pela Automação: x_string=" & x_string)
        MsgBox "Não foi possível ler o preço.", vbInformation, "Bico " & pBico & " em uso!"
    ElseIf x_quant = 16 Then
        x_string = MSCommAutomacao.Input
        Call CriaLogCupom(Time & " - Leitura de Preço pela Automação: x_string=" & x_string)
        'MsgBox "Bico: " & Mid(x_string, 4, 2) & Chr(10) & "Preco: " & Mid(x_string, 10, 1) & "," & Mid(x_string, 11, 3), vbInformation, "Leitura de Preço!"
        msk_venda.Text = Mid(x_string, 10, 1) & "," & Mid(x_string, 11, 3)
        'lAutomacaoAbertura = Mid(x_string, 6, 6) & "," & Mid(x_string, 12, 2)
        'lAutomacaoEncerrante = Mid(x_string, 6, 6) & "," & Mid(x_string, 12, 2)
        'AutomacaoAtualTabe
        lAutomacaoFlag = 0
        TimerAutomacao.Enabled = False
    End If
End Sub
Public Sub AutomacaoEnviaComandoPrecoAtual(ByVal pBico As String)
    Dim i As Integer
    Dim x_quant As Integer
    Dim x_string As String
    'Pede Preço Atual
    MSCommAutomacao.InBufferCount = 0
    x_string = "(&T" & pBico & "U"
    x_quant = 0
    For i = 2 To Len(x_string)
        x_quant = x_quant + Asc(Mid(x_string, i, 1))
        If x_quant > 255 Then
            x_quant = x_quant - 256
        End If
    Next
    If Len(Hex(x_quant)) = 1 Then
        x_string = x_string & "0" & Hex(x_quant) & ")"
    Else
        x_string = x_string & Hex(x_quant) & ")"
    End If
    Call CriaLogCupom(Time & " - AutomacaoEnviaComandoPrecoAtual: x_string=" & x_string)
    MSCommAutomacao.Output = x_string
    lAutomacaoFlag = 1
    TimerAutomacao.Enabled = True
End Sub
Public Sub AutomacaoEnviaComandoPrecoNovo(ByVal pBico As String)
    Dim i As Integer
    Dim x_quant As Integer
    Dim x_string As String
    Dim x_preco As String
    
    'Envia Preço Novo
    x_preco = Mid(Format(fValidaValor4(msk_venda), "0.000"), 1, 1) & Mid(Format(fValidaValor4(msk_venda), "0.000"), 3, 3)
    '(&UbbN0vvvvKK)
    MSCommAutomacao.InBufferCount = 0
    x_string = "(&U" & pBico & "10" & x_preco
    x_quant = 0
    For i = 2 To Len(x_string)
        x_quant = x_quant + Asc(Mid(x_string, i, 1))
        If x_quant > 256 Then
            x_quant = x_quant - 256
        End If
    Next
    If Len(Hex(x_quant)) = 1 Then
        x_string = x_string & "0" & Hex(x_quant) & ")"
    Else
        x_string = x_string & Hex(x_quant) & ")"
    End If
    Call CriaLogCupom(Time & " - AutomacaoEnviaComandoPrecoNovo: String=" & x_string)
    MSCommAutomacao.Output = x_string
    MsgBox "Comando Enviado"
    lAutomacaoFlag = 0
    TimerAutomacao.Enabled = False
End Sub
Private Sub DesativaBotoes()
    cmd_novo.Enabled = False
    cmd_alterar.Enabled = False
    cmd_excluir.Enabled = False
    cmd_sair.Enabled = False
    frm_move.Visible = False
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
End Sub
Private Sub Finaliza()
    Dim retval As Long
    If g_automacao Then
        If MSCommAutomacao.PortOpen = True Then
            MSCommAutomacao.PortOpen = False
        End If
        retval = Shell("C:\Cerrado\AutoCerrado\AutoCerrado.exe", vbMinimizedFocus)
    End If
    Set BicoAutomacao = Nothing
    Set Bomba = Nothing
    frm_cadastro.Show
End Sub
Private Sub Inclui()
    lOpcao = 1
    DesativaBotoes
    cmd_novo.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    txt_codigo.Text = 1
    If Bomba.LocalizarUltimo(g_empresa) Then
        txt_codigo.Text = Bomba.Codigo + 1
    End If
End Sub
Private Sub cbo_tipo_preco_GotFocus()
    SendMessageLong cbo_tipo_preco.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_tipo_preco_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dtcbo_produto.SetFocus
    End If
End Sub
Private Sub cmd_alterar_Click()
    lOpcao = 2
    DesativaBotoes
    cmd_alterar.Visible = True
    cmd_alterar.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    frm_dados.Enabled = True
    txt_codigo.Enabled = False
    msk_custo.SetFocus
End Sub
Private Sub cmd_anterior_Click()
    If Bomba.LocalizarAnterior Then
        AtualTela
    Else
        MsgBox "Início de Arquivo.", vbInformation, "Atenção!"
        cmd_proximo.SetFocus
    End If
End Sub
Private Sub cmd_cancelar_Click()
    LimpaTela
    lOpcao = 0
    If Bomba.LocalizarCodigo(g_empresa, lCodigo) Then
        AtualTela
        AtivaBotoes
        cmd_alterar.SetFocus
    Else
        DesativaBotoes
        cmd_novo.Enabled = True
        cmd_sair.Enabled = True
        cmd_novo.SetFocus
    End If
End Sub
Private Sub LimpaTela()
    txt_codigo.Text = ""
    dtcbo_combustivel.BoundText = ""
    txtNumeroSerie.Text = ""
    msk_custo.Text = ""
    msk_venda.Text = ""
    txt_tanque.Text = ""
    txt_numero_ilha.Text = ""
    cbo_tipo_preco.ListIndex = -1
    dtcbo_produto.BoundText = ""
End Sub
Private Sub cmd_excluir_Click()
    If Val(txt_codigo.Text) > 0 Then
        If MsgBox("Deseja excluir este registro?", vbQuestion + vbYesNo + vbDefaultButton2, "Exclusão de Registro!") = vbYes Then
            If Bomba.Excluir(g_empresa, Val(txt_codigo.Text)) Then
                LimpaTela
                If Bomba.LocalizarUltimo(g_empresa) Then
                    AtualTela
                Else
                    DesativaBotoes
                    cmd_novo.Enabled = True
                    cmd_sair.Enabled = True
                    cmd_novo.SetFocus
                End If
            Else
                MsgBox "Não foi possivel excluir este registro!", vbInformation, "Erro de Verificação!"
            End If
        End If
    End If
End Sub
Private Sub cmd_novo_Click()
'    Dim retval
'    SendKeys "^p" ', True
'    SendKeys "{ENTER}"
'    retval = Shell("D:\Msoffice.97\OFFICE\WINWORD.EXE E:\DOC\TABELA.DOC", 0)
    LimpaTela
    Inclui
    frm_dados.Enabled = True
    dtcbo_combustivel.SetFocus
End Sub
Private Sub cmd_ok_Click()
    On Error GoTo FileError
    If ValidaCampos Then
        AtivaBotoes
        If lOpcao = 1 Then
            AtualTabe
            If Bomba.Incluir Then
                lCodigo = Val(txt_codigo.Text)
                lPrecoVenda = fValidaValor4(msk_venda.Text)
            Else
                MsgBox "Não foi possível incluir este registro!", vbInformation, "Erro de Verificação!"
            End If
        ElseIf lOpcao = 2 Then
            If g_automacao Then
                If lPrecoVenda <> fValidaValor4(msk_venda.Text) Then
                    If MsgBox("Deseja mudar o preço pela automação?", vbQuestion + vbYesNo + vbDefaultButton2, "Alteração de Preço!") = vbYes Then
                        Call CriaLogCupom(Time & " - Alteração de Preço pela Automação: Bico=" & lCodigo & " Preço Anterior=" & Format(lPrecoVenda, "##,##0.000") & " Preço Novo=" & msk_venda.Text)
                        Call AutomacaoEnviaComandoPrecoNovo(ConverteBico(CStr(Format(lCodigo, "00"))))
                    End If
                End If
            End If
            AtualTabe
            If Bomba.Alterar(g_empresa, lCodigo) Then
                lCodigo = Val(txt_codigo.Text)
                lPrecoVenda = fValidaValor4(msk_venda.Text)
            Else
                MsgBox "Não foi possível alterar este registro!", vbInformation, "Erro de Verificação!"
            End If
        End If
        lOpcao = 0
        Call Bomba.LocalizarCodigo(g_empresa, lCodigo)
        AtualTela
        cmd_novo.SetFocus
    End If
    Exit Sub
FileError:
    MsgBox Error
    'ErroArquivo tbl_bomba.Name, "Bombaa"
    Exit Sub
End Sub
Function PortaDaAutomacao() As Integer
    
    On Error GoTo FileError
    
    PortaDaAutomacao = 2
    PortaDaAutomacao = ReadINI("CUPOM FISCAL", "Porta da Automacao", ArqSgpIni)
    Exit Function

FileError:
    Exit Function
End Function
Private Sub PreencheCboTipoPreco()
    cbo_tipo_preco.Clear
    cbo_tipo_preco.AddItem "À Vista"
    cbo_tipo_preco.ItemData(cbo_tipo_preco.NewIndex) = 0
    cbo_tipo_preco.AddItem "À Prazo"
    cbo_tipo_preco.ItemData(cbo_tipo_preco.NewIndex) = 1
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If Not Val(txt_codigo.Text) > 0 Then
        MsgBox "Informe o codigo do Bico.", vbInformation, "Atenção!"
        txt_codigo.SetFocus
    ElseIf IsNull(dtcbo_combustivel.SelectedItem) Then
        MsgBox "Escolha o Combustível.", vbInformation, "Atenção!"
        dtcbo_combustivel.SetFocus
    ElseIf txtNumeroSerie.Text = "" Then
        MsgBox "Informe o número de série da bomba.", vbInformation, "Atenção!"
        txtNumeroSerie.SetFocus
    ElseIf Not fValidaValor4(msk_custo.Text) > 0 Then
        MsgBox "Informe o Preço de Custo.", vbInformation, "Atenção!"
        msk_custo.SetFocus
    ElseIf Not fValidaValor4(msk_venda.Text) > 0 Then
        MsgBox "Informe o Preço de Venda.", vbInformation, "Atenção!"
        msk_venda.SetFocus
    ElseIf txt_tanque.Text = "" Then
        MsgBox "Informe o Número do Tanque.", vbInformation, "Atenção!"
        txt_tanque.SetFocus
    ElseIf Not Val(txt_numero_ilha.Text) > 0 Then
        MsgBox "O número da ilha deve ser maior que 0.", vbInformation, "Atenção!"
        txt_numero_ilha.SetFocus
    ElseIf cbo_tipo_preco.ListIndex = -1 Then
        MsgBox "Escolha um tipo de preço.", vbInformation, "Atenção!"
        cbo_tipo_preco.SetFocus
    ElseIf IsNull(dtcbo_produto.SelectedItem) Then
        MsgBox "Escolha um produto.", vbInformation, "Atenção!"
        dtcbo_produto.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_primeiro_Click()
    If Bomba.LocalizarPrimeiro(g_empresa) Then
        AtualTela
        cmd_proximo.SetFocus
    Else
        MsgBox "Não há registro.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    If Bomba.LocalizarProximo Then
        AtualTela
    Else
        MsgBox "Fim de Arquivo.", vbInformation, "Atenção!"
        cmd_anterior.SetFocus
    End If
End Sub
Private Sub cmd_sair_Click()
'    SendKeys "%{F4}"
'    AppActivate MyAppID ' Activate Microsoft
    Unload Me
End Sub
Private Sub cmd_ultimo_Click()
    If Bomba.LocalizarUltimo(g_empresa) Then
        AtualTela
        cmd_anterior.SetFocus
    Else
        MsgBox "Não há registro.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub dtcbo_combustivel_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        txtNumeroSerie.SetFocus
    End If
End Sub
Private Sub dtcbo_produto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub Form_Activate()
    If lFlagCadastro = 0 Then
        DesativaBotoes
        If Bomba.LocalizarUltimo(g_empresa) Then
            AtualTela
            AtivaBotoes
        Else
            cmd_novo.Enabled = True
            cmd_sair.Enabled = True
        End If
        cmd_novo.SetFocus
    Else
        lFlagCadastro = 0
    End If
    If g_automacao Then
        '----- Porta
        lAutomacaoPorta = PortaDaAutomacao
        '----- Velocidade
        lAutomacaoVelocidade = "9600,n,8,1"
        '----- DTR
        lAutomacaoDtr = True
        'ldtr = False
        '----- RTS
        lAutomacaoRts = True
        'lRts = False
        '----- EOF
        AutomacaoAbrePorta
        lAutomacaoFlag = 0
        lAutomacaoBicoEncerrante = 1
    End If
    TimerAutomacao.Enabled = False
End Sub
Private Sub Form_Deactivate()
    lFlagCadastro = 1
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 And lOpcao = 0 Then
        KeyCode = 0
        cmd_novo_Click
    ElseIf KeyCode = vbKeyF3 And lOpcao = 0 Then
        KeyCode = 0
        cmd_alterar_Click
    ElseIf KeyCode = vbKeyF4 And Shift = 0 And lOpcao = 0 Then
        KeyCode = 0
        cmd_excluir_Click
    ElseIf KeyCode = vbKeyF7 And lOpcao = 0 Then
        KeyCode = 0
        cmd_primeiro_Click
    ElseIf KeyCode = vbKeyF8 And lOpcao = 0 Then
        KeyCode = 0
        cmd_anterior_Click
    ElseIf KeyCode = vbKeyF9 And lOpcao = 0 Then
        KeyCode = 0
        cmd_proximo_Click
    ElseIf KeyCode = vbKeyF10 And lOpcao = 0 Then
        KeyCode = 0
        cmd_ultimo_Click
    ElseIf KeyCode = vbKeyF11 And lOpcao > 0 Then
        KeyCode = 0
        cmd_ok_Click
    ElseIf KeyCode = vbKeyF12 And lOpcao > 0 Then
        KeyCode = 0
        cmd_cancelar_Click
    End If
End Sub
Private Sub Form_Load()
    Screen.MousePointer = 1
    CentraForm Me
    If ComunicaAutomacaoCerrado("AUTOMACAO PAUSA", "") Then
        If ComunicaAutomacaoCerrado("AUTOMACAO FECHA", "") Then
        End If
    Else
    End If
    
    
    '''teste pra "resolver" problema da serial
    g_empresa = 1
    g_automacao = True
    'g_nome_empresa = "POSTO DO BOSQUE M. Be LTDA"



'    adodc_combustivel.ConnectionString = gConnectionString
'    adodc_combustivel.RecordSource = "SELECT Codigo, Nome FROM Combustivel WHERE Empresa = " & g_empresa & " ORDER BY Nome"
'    adodc_combustivel.Refresh
    Set adodc_combustivel.Recordset = Conectar.RsConexao("SELECT Codigo, Nome FROM Combustivel WHERE Empresa = " & g_empresa & " ORDER BY Nome")
'    adodc_produto.ConnectionString = gConnectionString
'    adodc_produto.RecordSource = "SELECT Codigo, Nome FROM Produto WHERE Inativo = FALSE AND [Codigo do Grupo] = 4 ORDER BY Nome"
'    adodc_produto.Refresh
    Set adodc_produto.Recordset = Conectar.RsConexao("SELECT Codigo, Nome FROM Produto WHERE Inativo = " & preparaBooleano(False) & " AND [Codigo do Grupo] = 4 ORDER BY Nome")
    PreencheCboTipoPreco
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_custo_GotFocus()
    msk_custo.SelStart = 0
    msk_custo.SelLength = Len(msk_custo.Text)
End Sub
Private Sub msk_custo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        msk_venda.SetFocus
    End If
End Sub
Private Sub msk_custo_LostFocus()
    msk_custo.Text = Format(msk_custo.Text, "###,##0.0000")
End Sub
Private Sub msk_venda_GotFocus()
    If g_automacao Then
        If lOpcao = 2 Then
            Call AutomacaoEnviaComandoPrecoAtual(ConverteBico(CStr(Format(lCodigo, "00"))))
        End If
    End If
    msk_venda.SelStart = 0
    msk_venda.SelLength = Len(msk_venda.Text)
End Sub
Private Sub msk_venda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_tanque.SetFocus
    End If
End Sub
Private Sub msk_venda_LostFocus()
    msk_venda.Text = Format(msk_venda.Text, "###,##0.0000")
End Sub
Private Sub TimerAutomacao_Timer()
    'Pede Preco - lAutomacaoFlag = 1
    ' (&U0300143208) $1,432 bico 03 nivel a vista
    If lAutomacaoFlag = 1 Then
        Call AutomacaoBuscaPreco(ConverteBico(CStr(Format(lCodigo, "00"))))
        lAutomacaoFlag = 0
    End If
End Sub
Private Sub txt_codigo_GotFocus()
    txt_codigo.SelStart = 0
    txt_codigo.SelLength = Len(txt_codigo)
End Sub
Private Sub txt_codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dtcbo_combustivel.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_codigo_LostFocus()
    If lOpcao = 1 And txt_codigo <> "" Then
        If Bomba.LocalizarCodigo(g_empresa, Val(txt_codigo.Text)) Then
            MsgBox "Já existe bico de combustível cadastrado com este código." & Chr(10) & Chr(10) & "Mude o código informado.", vbInformation, "Duplicidade de Registro!"
            txt_codigo.SetFocus
            Exit Sub
        End If
    End If
End Sub
Private Sub txt_numero_ilha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_tipo_preco.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_tanque_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_numero_ilha.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txtNumeroSerie_GotFocus()
    txtNumeroSerie.SelStart = 0
    txtNumeroSerie.SelLength = Len(txtNumeroSerie.Text)
End Sub
Private Sub txtNumeroSerie_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_custo.SetFocus
    End If
End Sub
