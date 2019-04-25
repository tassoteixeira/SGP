VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form baixa_nota_abastecimento_periodo 
   Caption         =   "Baixa de Notas de Abastecimento p/ Período"
   ClientHeight    =   2475
   ClientLeft      =   1920
   ClientTop       =   2790
   ClientWidth     =   5175
   Icon            =   "baixa_nota_abastecimento_periodo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "baixa_nota_abastecimento_periodo.frx":030A
   ScaleHeight     =   2475
   ScaleWidth      =   5175
   Begin VB.CheckBox chkNotaConferida 
      Caption         =   "&Notas Conferidas"
      Height          =   255
      Left            =   2820
      TabIndex        =   1
      Top             =   120
      Width           =   2235
   End
   Begin VB.CommandButton cmd_baixa 
      Caption         =   "&Baixar"
      Height          =   855
      Left            =   180
      Picture         =   "baixa_nota_abastecimento_periodo.frx":0750
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Baixa as notas de abastecimento no período informado."
      Top             =   1500
      Width           =   795
   End
   Begin VB.CommandButton cmd_extornar 
      Caption         =   "&Extornar"
      Height          =   855
      Left            =   2160
      Picture         =   "baixa_nota_abastecimento_periodo.frx":1A2A
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Extorna as notas de abastecimento baixadas no período informado."
      Top             =   1500
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4200
      Picture         =   "baixa_nota_abastecimento_periodo.frx":2D04
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   1500
      Width           =   795
   End
   Begin MSMask.MaskEdBox msk_data_baixa 
      Height          =   300
      Left            =   3960
      TabIndex        =   8
      Top             =   1080
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox msk_data_final 
      Height          =   300
      Left            =   2040
      TabIndex        =   6
      Top             =   1080
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox msk_data_inicial 
      Height          =   300
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSAdodcLib.Adodc adodcCliente 
      Height          =   330
      Left            =   2040
      Top             =   420
      Visible         =   0   'False
      Width           =   2475
      _ExtentX        =   4366
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
      Caption         =   "adodcCliente"
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
   Begin MSDataListLib.DataCombo dtcboCliente 
      Bindings        =   "baixa_nota_abastecimento_periodo.frx":4396
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   420
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "Razao Social"
      BoundColumn     =   "Codigo"
      Text            =   "dtcboCliente"
   End
   Begin VB.Label Label1 
      Caption         =   "Da&ta da Baixa"
      Height          =   255
      Index           =   3
      Left            =   3960
      TabIndex        =   7
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "D&ata Final"
      Height          =   255
      Index           =   2
      Left            =   2040
      TabIndex        =   5
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "&Data Inicial"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "&Cliente"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "baixa_nota_abastecimento_periodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lNumeroMovimentoCaixa As Long
Dim lClienteAnterior As Long
Dim lNumeroDocumento As Long
Dim lTotal As Currency

Private BaixaDuplicataReceber As New cBaixaDuplicataReceber
Private Cliente As New cCliente
Private ConfiguracaoDiversa As New cConfiguracaoDiversa
Private IntegracaoCaixa As New cIntegracaoCaixa
Private MovBaixaNotaAbastecimento As New cBaixaNotaAbastecimento
Private MovCaixa As New cMovimentoCaixa
Private MovDuplicataReceber As New cDuplicataReceber
Private MovNotaAbastecimento As New cMovimentoNotaAbastecimento
Private Sub PreparaBaixa()
    Dim xString As String
    Dim xConferencia As Boolean
    
    If chkNotaConferida.Value = 0 Then
        xConferencia = False
    Else
        xConferencia = True
    End If
    'lTotal = MovNotaAbastecimento.TotalDataLiquido(g_empresa, CLng(dtcboCliente.BoundText), CDate(msk_data_inicial.Text), CDate(msk_data_final.Text), xConferencia)
    lTotal = MovNotaAbastecimento.TotalDataLiquido(0, CLng(dtcboCliente.BoundText), CDate(msk_data_inicial.Text), CDate(msk_data_final.Text), xConferencia)
    If lTotal > 0 Then
        If (MsgBox("Este cliente tem R$ " & Format(lTotal, "###,###,##0.00") & " de notas no peíodo informado." & Chr(13) & Chr(10) & "Deseja realmente baixá-las?", vbQuestion + vbYesNo + vbDefaultButton2, "Baixa de Notas de Abastecimento.")) = vbYes Then
            xString = xString & "Cli:" & CLng(dtcboCliente.BoundText)
            xString = xString & " Ref:" & msk_data_inicial.Text & " a " & msk_data_final.Text
            xString = xString & " Vlr:" & Format(lTotal, "###,###,##0.00")
            Call GravaAuditoria(1, Me.name, 18, xString)
            Baixa
        End If
    Else
        MsgBox "Este cliente não tem notas de abastecimento neste período!", 48, "Baixa de Notas de Abastecimento."
        dtcboCliente.SetFocus
    End If
End Sub
Private Sub PreparaExtorno()
    Dim xTotal As Currency
    Dim xString As String
    Dim xConferencia As Boolean
    
    If chkNotaConferida.Value = 0 Then
        xConferencia = False
    Else
        xConferencia = True
    End If
    'xTotal = MovBaixaNotaAbastecimento.TotalDataLiquido(g_empresa, CLng(dtcboCliente.BoundText), CDate(msk_data_inicial.Text), CDate(msk_data_final.Text), xConferencia)
    xTotal = MovBaixaNotaAbastecimento.TotalDataLiquido(0, CLng(dtcboCliente.BoundText), CDate(msk_data_inicial.Text), CDate(msk_data_final.Text), xConferencia)
    If xTotal > 0 Then
        If (MsgBox("Este cliente tem R$ " & Format(xTotal, "###,###,##0.00") & " de notas baixadas no peíodo informado." & Chr(13) & Chr(10) & "Deseja realmente extorna-las?", 4 + 32 + 256, "Extorno de Notas de Abastecimento.")) = vbYes Then
            xString = xString & "Cli:" & CLng(dtcboCliente.BoundText)
            xString = xString & " Ref:" & msk_data_inicial.Text & " a " & msk_data_final.Text
            xString = xString & " Vlr:" & Format(xTotal, "###,###,##0.00")
            Call GravaAuditoria(1, Me.name, 19, xString)
            Extorno
        End If
    Else
        MsgBox "Este cliente não tem notas de abastecimento baixadas neste período!", 48, "Extorno de Notas de Abastecimento."
        dtcboCliente.SetFocus
    End If
End Sub
Private Sub Baixa()
    Dim rsNotaAbastecimento As New adodb.Recordset
    Dim xSQL As String
    
    On Error GoTo FileError
    
    If (MsgBox("Deseja gerar Duplicata Baixada?" & vbCrLf & "Escolha SIM - para gerar Duplicata à Receber já Baixada." & vbCrLf & "Escolha NÃO - para gerar Duplicata à Receber.", vbQuestion + vbYesNo + vbDefaultButton1, "Tipo de Duplicata!")) = vbYes Then
        If Not CriaBaixaDuplicataReceber Then
            MsgBox "Não foi possível fazer esta baixa!", vbCritical, "Erro de Integridade."
            Exit Sub
        End If
    Else
        If Not CriaDuplicataAReceber Then
            MsgBox "Não foi possível fazer esta baixa!", vbCritical, "Erro de Integridade."
            Exit Sub
        End If
    End If
    
    xSQL = ""
    xSQL = xSQL & "SELECT [Codigo do Cliente], [Data do Abastecimento], [Numero da Nota], Ordem, [Codigo do Produto2], [Periodo], Empresa"
    xSQL = xSQL & "  FROM Movimento_Nota_Abastecimento"
    'xSQL = xSQL & " WHERE Empresa = " & g_empresa
    xSQL = xSQL & " WHERE Empresa >= " & 1
    xSQL = xSQL & "   AND [Codigo do Cliente] = " & CLng(dtcboCliente.BoundText)
    If chkNotaConferida.Value = 0 Then
        xSQL = xSQL & "   AND [Data do Abastecimento] >= " & preparaData(CDate(msk_data_inicial.Text))
        xSQL = xSQL & "   AND [Data do Abastecimento] <= " & preparaData(CDate(msk_data_final.Text))
    Else
        xSQL = xSQL & "   AND [Data da Conferencia] >= " & preparaData(CDate(msk_data_inicial.Text))
        xSQL = xSQL & "   AND [Data da Conferencia] <= " & preparaData(CDate(msk_data_final.Text))
    End If
    xSQL = xSQL & " ORDER BY [Data do Abastecimento], [Periodo], [Numero da Nota], Ordem, [Codigo do Produto2]"
    Set rsNotaAbastecimento = Conectar.RsConexao(xSQL)
    With rsNotaAbastecimento
        If .RecordCount > 0 Then
            Do Until .EOF
                'If MovNotaAbastecimento.LocalizarCodigo(g_empresa, rsNotaAbastecimento("Codigo do Cliente").Value, rsNotaAbastecimento("Data do Abastecimento").Value, rsNotaAbastecimento("Numero da Nota").Value, rsNotaAbastecimento("Ordem").Value, rsNotaAbastecimento("Codigo do Produto2").Value, rsNotaAbastecimento("Periodo").Value) Then
                If MovNotaAbastecimento.LocalizarCodigo(rsNotaAbastecimento("Empresa").Value, rsNotaAbastecimento("Codigo do Cliente").Value, rsNotaAbastecimento("Data do Abastecimento").Value, rsNotaAbastecimento("Numero da Nota").Value, rsNotaAbastecimento("Ordem").Value, rsNotaAbastecimento("Codigo do Produto2").Value, rsNotaAbastecimento("Periodo").Value) Then
                    If MovNotaAbastecimento.BaixadoPelaDuplicata = False Then
                        If Not IncluiMovimentoCaixa Then
                            MsgBox "Não foi possível integrar com o Caixa!", vbInformation, "Erro de Integridade."
                        End If
                    Else
                        lNumeroMovimentoCaixa = 0
                    End If
                    MovBaixaNotaAbastecimento.CodigoCliente = MovNotaAbastecimento.CodigoCliente
                    MovBaixaNotaAbastecimento.DataPagamento = msk_data_baixa.Text
                    MovBaixaNotaAbastecimento.CodigoProduto2 = MovNotaAbastecimento.CodigoProduto2
                    MovBaixaNotaAbastecimento.NumeroNota = MovNotaAbastecimento.NumeroNota
                    MovBaixaNotaAbastecimento.Ordem = MovNotaAbastecimento.Ordem
                    MovBaixaNotaAbastecimento.Empresa = MovNotaAbastecimento.Empresa
                    MovBaixaNotaAbastecimento.Periodo = MovNotaAbastecimento.Periodo
                    MovBaixaNotaAbastecimento.DataAbastecimento = MovNotaAbastecimento.DataAbastecimento
                    MovBaixaNotaAbastecimento.Quantidade = MovNotaAbastecimento.Quantidade
                    MovBaixaNotaAbastecimento.ValorUnitario = MovNotaAbastecimento.ValorUnitario
                    MovBaixaNotaAbastecimento.ValorTotal = MovNotaAbastecimento.ValorTotal
                    MovBaixaNotaAbastecimento.CodigoConveniado = MovNotaAbastecimento.CodigoConveniado
                    MovBaixaNotaAbastecimento.TipoMovimento = MovNotaAbastecimento.TipoMovimento
                    MovBaixaNotaAbastecimento.PlacaLetra = MovNotaAbastecimento.PlacaLetra
                    MovBaixaNotaAbastecimento.PlacaNumero = MovNotaAbastecimento.PlacaNumero
                    MovBaixaNotaAbastecimento.Historico = MovNotaAbastecimento.Historico
                    MovBaixaNotaAbastecimento.ValorPago = MovNotaAbastecimento.ValorTotal
                    MovBaixaNotaAbastecimento.ValorDescontoUnitario = MovNotaAbastecimento.ValorDescontoUnitario
                    MovBaixaNotaAbastecimento.NumeroMovimentoCaixa = MovNotaAbastecimento.NumeroMovimentoCaixa
                    MovBaixaNotaAbastecimento.NumeroMovimentoCaixaBaixa = lNumeroMovimentoCaixa
                    MovBaixaNotaAbastecimento.BaixadoPelaDuplicata = MovNotaAbastecimento.BaixadoPelaDuplicata
                    MovBaixaNotaAbastecimento.NumeroIlha = MovNotaAbastecimento.NumeroIlha
                    MovBaixaNotaAbastecimento.Origem = MovNotaAbastecimento.Origem
                    MovBaixaNotaAbastecimento.DataBaixa = msk_data_baixa.Text
                    MovBaixaNotaAbastecimento.NumeroCupom = MovNotaAbastecimento.NumeroCupom
                    MovBaixaNotaAbastecimento.DataConferencia = MovNotaAbastecimento.DataConferencia
                    MovBaixaNotaAbastecimento.NumeroDuplicata = lNumeroDocumento
                    MovBaixaNotaAbastecimento.KM = MovNotaAbastecimento.KM
                    If MovBaixaNotaAbastecimento.Incluir Then
                        If Not MovNotaAbastecimento.Excluir(MovNotaAbastecimento.Empresa, MovNotaAbastecimento.CodigoCliente, MovNotaAbastecimento.DataAbastecimento, MovNotaAbastecimento.NumeroNota, MovNotaAbastecimento.Ordem, MovNotaAbastecimento.CodigoProduto2, MovNotaAbastecimento.Periodo) Then
                            MsgBox "Não foi possível excluir este registro!", vbInformation, "Erro de Integridade."
                        End If
                    Else
                        MsgBox "Não foi possível incluir este registro!", vbInformation, "Erro de Integridade."
                    End If
                Else
                    MsgBox "Não foi possível localizar este registro!", vbInformation, "Erro de Integridade."
                End If
                .MoveNext
            Loop
        End If
    End With
    rsNotaAbastecimento.Close
    Set rsNotaAbastecimento = Nothing
    MsgBox "Baixa de notas de abastecimento concluida!", 48, "Fim da Baixa de Notas de Abastecimento."
    dtcboCliente.SetFocus
    Exit Sub

FileError:
    'ErroArquivo tbl_baixa_nota_abastecimento.Name, "Baixaa"
    Exit Sub
End Sub
Private Function CriaBaixaDuplicataReceber() As Boolean
    Dim xDataInicial As Date
    Dim xDataFinal As Date
    Dim xSQL As String
    Dim rsNotaAbastecimento As New adodb.Recordset
    
    CriaBaixaDuplicataReceber = False
    
    xDataInicial = CDate(msk_data_inicial.Text)
    xDataFinal = CDate(msk_data_final.Text)
    
    'Busa Data Inicial e Data Final
    If chkNotaConferida.Value = 1 Then
        xSQL = ""
        xSQL = xSQL & "SELECT [Data do Abastecimento]"
        xSQL = xSQL & "  FROM Movimento_Nota_Abastecimento"
        'xSQL = xSQL & " WHERE Empresa = " & g_empresa
        xSQL = xSQL & " WHERE Empresa >= " & 1
        xSQL = xSQL & "   AND [Codigo do Cliente] = " & CLng(dtcboCliente.BoundText)
        xSQL = xSQL & "   AND [Data da Conferencia] >= " & preparaData(CDate(msk_data_inicial.Text))
        xSQL = xSQL & "   AND [Data da Conferencia] <= " & preparaData(CDate(msk_data_final.Text))
        xSQL = xSQL & " ORDER BY [Data do Abastecimento]"
        Set rsNotaAbastecimento = Conectar.RsConexao(xSQL)
        If rsNotaAbastecimento.RecordCount > 0 Then
            rsNotaAbastecimento.MoveFirst
            xDataInicial = rsNotaAbastecimento("Data do Abastecimento").Value
            rsNotaAbastecimento.MoveLast
            xDataFinal = rsNotaAbastecimento("Data do Abastecimento").Value
        End If
        rsNotaAbastecimento.Close
    End If
    Set rsNotaAbastecimento = Nothing
    
    lNumeroDocumento = MovDuplicataReceber.ProximoNumeroDocumento(g_empresa)
    BaixaDuplicataReceber.Empresa = Val(g_empresa)
    BaixaDuplicataReceber.NumeroDocumento = lNumeroDocumento
    BaixaDuplicataReceber.CodigoCliente = CLng(dtcboCliente.BoundText)
    BaixaDuplicataReceber.DataEmissao = msk_data_baixa.Text
    BaixaDuplicataReceber.DataPeriodoInicial = xDataInicial
    BaixaDuplicataReceber.DataPeriodoFinal = xDataFinal
    BaixaDuplicataReceber.NumeroDuplicata = "1"
    BaixaDuplicataReceber.DataVencimento = msk_data_baixa.Text
    BaixaDuplicataReceber.ValorVencimento = lTotal
    BaixaDuplicataReceber.ValorCustoBancario = 0
    BaixaDuplicataReceber.DataPagamento = msk_data_baixa.Text
    BaixaDuplicataReceber.ValorJuros = 0
    BaixaDuplicataReceber.ValorPago = lTotal
    BaixaDuplicataReceber.ValorPagoChequeVista = 0
    BaixaDuplicataReceber.ValorPagoChequePrazo = 0
    BaixaDuplicataReceber.Periodo = 1
    BaixaDuplicataReceber.NumeroMovimentoCaixa = 0
    BaixaDuplicataReceber.NumeroMovimentoCaixaBaixa = 0
    BaixaDuplicataReceber.ValorDesconto = 0
    BaixaDuplicataReceber.ValorRestante = 0
    BaixaDuplicataReceber.DataVencimentoRestante = "00:00:00"
    If Not BaixaDuplicataReceber.Incluir Then
        MsgBox "Não foi possível gerar Baixa de Duplicata à Receber!", vbCritical, "Erro de Integridade."
    Else
        CriaBaixaDuplicataReceber = True
    End If
End Function
Private Function CriaDuplicataAReceber() As Boolean
    Dim xDataInicial As Date
    Dim xDataFinal As Date
    Dim xSQL As String
    Dim rsNotaAbastecimento As New adodb.Recordset
    
    CriaDuplicataAReceber = False
    
    xDataInicial = CDate(msk_data_inicial.Text)
    xDataFinal = CDate(msk_data_final.Text)
    
    'Busa Data Inicial e Data Final
    If chkNotaConferida.Value = 1 Then
        xSQL = ""
        xSQL = xSQL & "SELECT [Data do Abastecimento]"
        xSQL = xSQL & "  FROM Movimento_Nota_Abastecimento"
        xSQL = xSQL & " WHERE Empresa = " & g_empresa
        xSQL = xSQL & "   AND [Codigo do Cliente] = " & CLng(dtcboCliente.BoundText)
        xSQL = xSQL & "   AND [Data da Conferencia] >= " & preparaData(CDate(msk_data_inicial.Text))
        xSQL = xSQL & "   AND [Data da Conferencia] <= " & preparaData(CDate(msk_data_final.Text))
        xSQL = xSQL & " ORDER BY [Data do Abastecimento]"
        Set rsNotaAbastecimento = Conectar.RsConexao(xSQL)
        If rsNotaAbastecimento.RecordCount > 0 Then
            rsNotaAbastecimento.MoveFirst
            xDataInicial = rsNotaAbastecimento("Data do Abastecimento").Value
            rsNotaAbastecimento.MoveLast
            xDataFinal = rsNotaAbastecimento("Data do Abastecimento").Value
        End If
        rsNotaAbastecimento.Close
    End If
    Set rsNotaAbastecimento = Nothing
    
    lNumeroDocumento = MovDuplicataReceber.ProximoNumeroDocumento(g_empresa)
    MovDuplicataReceber.Empresa = Val(g_empresa)
    MovDuplicataReceber.NumeroDocumento = lNumeroDocumento
    MovDuplicataReceber.CodigoCliente = CLng(dtcboCliente.BoundText)
    MovDuplicataReceber.DataEmissao = msk_data_baixa.Text
    MovDuplicataReceber.DataPeriodoInicial = xDataInicial
    MovDuplicataReceber.DataPeriodoFinal = xDataFinal
    MovDuplicataReceber.NumeroDuplicata = "1"
    MovDuplicataReceber.DataVencimento = msk_data_baixa.Text
    MovDuplicataReceber.ValorVencimento = lTotal
    MovDuplicataReceber.ValorCustoBancario = 0
    MovDuplicataReceber.NumeroMovimentoCaixa = 0
    If Not MovDuplicataReceber.Incluir Then
        MsgBox "Não foi possível gerar Duplicata à Receber!", vbCritical, "Erro de Integridade."
    Else
        CriaDuplicataAReceber = True
    End If
End Function
Private Sub ExcluiMovimentoCaixa()
    If Not MovCaixa.Excluir(g_empresa, MovBaixaNotaAbastecimento.DataPagamento, MovBaixaNotaAbastecimento.NumeroMovimentoCaixaBaixa) Then
        MsgBox "Não foi excluído o movimento do caixa!", vbInformation, "Erro de Integridade."
    End If
End Sub
Private Sub Extorno()
    Dim rsBaixaNotaAbastecimento As New adodb.Recordset
    Dim xSQL As String
    
    On Error GoTo FileError
    
    xSQL = ""
    xSQL = xSQL & "SELECT [Codigo do Cliente], [Data do Abastecimento], [Numero da Nota], Ordem, [Codigo do Produto2], [Periodo], Empresa"
    xSQL = xSQL & "  FROM Baixa_Nota_Abastecimento"
    'xSQL = xSQL & " WHERE Empresa = " & g_empresa
    xSQL = xSQL & " WHERE Empresa >= " & 1
    xSQL = xSQL & "   AND [Codigo do Cliente] = " & CLng(dtcboCliente.BoundText)
    If chkNotaConferida.Value = 0 Then
        xSQL = xSQL & "   AND [Data do Abastecimento] >= " & preparaData(CDate(msk_data_inicial.Text))
        xSQL = xSQL & "   AND [Data do Abastecimento] <= " & preparaData(CDate(msk_data_final.Text))
    Else
        xSQL = xSQL & "   AND [Data da Conferencia] >= " & preparaData(CDate(msk_data_inicial.Text))
        xSQL = xSQL & "   AND [Data da Conferencia] <= " & preparaData(CDate(msk_data_final.Text))
    End If
    xSQL = xSQL & " ORDER BY [Data do Abastecimento], [Periodo], [Numero da Nota], Ordem, [Codigo do Produto2]"
    Set rsBaixaNotaAbastecimento = Conectar.RsConexao(xSQL)
    With rsBaixaNotaAbastecimento
        If .RecordCount > 0 Then
            Do Until .EOF
                If MovBaixaNotaAbastecimento.LocalizarCodigo(rsBaixaNotaAbastecimento("Empresa").Value, rsBaixaNotaAbastecimento("Codigo do Cliente").Value, rsBaixaNotaAbastecimento("Data do Abastecimento").Value, rsBaixaNotaAbastecimento("Numero da Nota").Value, rsBaixaNotaAbastecimento("Ordem").Value, rsBaixaNotaAbastecimento("Codigo do Produto2").Value, rsBaixaNotaAbastecimento("Periodo").Value) Then
                    If MovBaixaNotaAbastecimento.NumeroMovimentoCaixaBaixa > 0 Then
                        Call ExcluiMovimentoCaixa
                    End If
                    MovNotaAbastecimento.Empresa = MovBaixaNotaAbastecimento.Empresa
                    MovNotaAbastecimento.CodigoCliente = MovBaixaNotaAbastecimento.CodigoCliente
                    MovNotaAbastecimento.DataAbastecimento = MovBaixaNotaAbastecimento.DataAbastecimento
                    MovNotaAbastecimento.NumeroNota = MovBaixaNotaAbastecimento.NumeroNota
                    MovNotaAbastecimento.Ordem = MovBaixaNotaAbastecimento.Ordem
                    MovNotaAbastecimento.CodigoProduto2 = MovBaixaNotaAbastecimento.CodigoProduto2
                    MovNotaAbastecimento.Periodo = MovBaixaNotaAbastecimento.Periodo
                    MovNotaAbastecimento.Quantidade = MovBaixaNotaAbastecimento.Quantidade
                    MovNotaAbastecimento.ValorUnitario = MovBaixaNotaAbastecimento.ValorUnitario
                    MovNotaAbastecimento.ValorTotal = MovBaixaNotaAbastecimento.ValorTotal
                    MovNotaAbastecimento.CodigoConveniado = MovBaixaNotaAbastecimento.CodigoConveniado
                    MovNotaAbastecimento.TipoMovimento = MovBaixaNotaAbastecimento.TipoMovimento
                    MovNotaAbastecimento.PlacaLetra = MovBaixaNotaAbastecimento.PlacaLetra
                    MovNotaAbastecimento.PlacaNumero = MovBaixaNotaAbastecimento.PlacaNumero
                    MovNotaAbastecimento.Historico = MovBaixaNotaAbastecimento.Historico
                    MovNotaAbastecimento.ValorDescontoUnitario = MovBaixaNotaAbastecimento.ValorDescontoUnitario
                    MovNotaAbastecimento.NumeroMovimentoCaixa = MovBaixaNotaAbastecimento.NumeroMovimentoCaixa
                    MovNotaAbastecimento.BaixadoPelaDuplicata = MovBaixaNotaAbastecimento.BaixadoPelaDuplicata
                    MovNotaAbastecimento.NumeroIlha = MovBaixaNotaAbastecimento.NumeroIlha
                    MovNotaAbastecimento.Origem = MovBaixaNotaAbastecimento.Origem
                    MovNotaAbastecimento.NumeroCupom = MovBaixaNotaAbastecimento.NumeroCupom
                    MovNotaAbastecimento.DataConferencia = MovBaixaNotaAbastecimento.DataConferencia
                    MovNotaAbastecimento.KM = MovBaixaNotaAbastecimento.KM
                    If MovNotaAbastecimento.Incluir Then
                        If Not MovBaixaNotaAbastecimento.Excluir(MovBaixaNotaAbastecimento.Empresa, MovBaixaNotaAbastecimento.CodigoCliente, MovBaixaNotaAbastecimento.DataAbastecimento, MovBaixaNotaAbastecimento.NumeroNota, MovBaixaNotaAbastecimento.Ordem, MovBaixaNotaAbastecimento.CodigoProduto2, MovBaixaNotaAbastecimento.Periodo) Then
                            MsgBox "Não foi possível excluir este registro!", vbInformation, "Erro de Integridade."
                        End If
                    Else
                        MsgBox "Não foi possível incluir este registro!", vbInformation, "Erro de Integridade."
                    End If
                Else
                    MsgBox "Não foi possível localizar este registro!", vbInformation, "Erro de Integridade."
                End If
                .MoveNext
            Loop
        End If
    End With
    rsBaixaNotaAbastecimento.Close
    Set rsBaixaNotaAbastecimento = Nothing
    MsgBox "Extorno de notas de abastecimento concluido!", 48, "Fim do Extorno de Notas de Abastecimento."
    dtcboCliente.SetFocus
    Exit Sub

FileError:
    'ErroArquivo tbl_movimento_nota.Name, "Notaa"
    Exit Sub
End Sub
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    FinalizaProcessoCaixa
    Set BaixaDuplicataReceber = Nothing
    Set Cliente = Nothing
    Set ConfiguracaoDiversa = Nothing
    Set IntegracaoCaixa = Nothing
    Set MovBaixaNotaAbastecimento = Nothing
    Set MovCaixa = Nothing
    Set MovDuplicataReceber = Nothing
    Set MovNotaAbastecimento = Nothing
End Sub
Function IncluiMovimentoCaixa() As Boolean
    Dim xComplemento As String
    IncluiMovimentoCaixa = False
    lNumeroMovimentoCaixa = 0
    
    If IntegracaoCaixa.LocalizarNome(g_empresa, "DUPLICATAS A RECEBER") Then
        xComplemento = "TM:" & MovNotaAbastecimento.TipoMovimento & " P:" & MovNotaAbastecimento.Periodo & " " & Cliente.RazaoSocial
        MovCaixa.Empresa = g_empresa
        MovCaixa.Data = Format(msk_data_baixa.Text, "dd/mm/yyyy")
        MovCaixa.NumeroMovimento = 1
        MovCaixa.Valor = MovNotaAbastecimento.ValorTotal
        MovCaixa.NumeroDocumento = MovNotaAbastecimento.NumeroNota
        MovCaixa.CodigoHistorico = IntegracaoCaixa.HistoricoPadrao
        MovCaixa.Complemento = xComplemento
        MovCaixa.NumeroContaDebito = IntegracaoCaixa.ContaDebito
        MovCaixa.NumeroContaCredito = IntegracaoCaixa.ContaCredito
        MovCaixa.TipoMovimento = 2
        MovCaixa.FluxoCaixa = True
        MovCaixa.CodigoUsuario = g_usuario
        If MovCaixa.Incluir > 0 Then
            IncluiMovimentoCaixa = True
            lNumeroMovimentoCaixa = MovCaixa.NumeroMovimento
        Else
            MsgBox "Não foi integrado no caixa o valor=" & MovNotaAbastecimento.ValorTotal, vbInformation, "Erro de Integridade"
        End If
    Else
        MsgBox "Não existe a integração=" & "DUPLICATAS A RECEBER" & ".", vbInformation, "Registro Inexistente"
    End If
End Function
Private Sub cmd_baixa_Click()
    If ValidaCampos Then
        PreparaBaixa
    End If
End Sub
Private Sub cmd_extornar_Click()
    If ValidaCampos Then
'        Shell "MPLAY32.EXE C:\VB5\GRAPHICS\AVIs\FILEDELR.AVI", vbMinimizedFocus
        PreparaExtorno
    End If
End Sub
Function ValidaCampos() As Boolean
    ValidaCampos = False
    If Val(dtcboCliente.BoundText) = 0 Then
        MsgBox "Informe o cliente.", vbInformation, "Atenção!"
        dtcboCliente.SetFocus
    ElseIf Not IsDate(msk_data_baixa.Text) Then
        MsgBox "Informe a data da baixa.", vbInformation, "Atenção!"
        msk_data_baixa.SetFocus
    ElseIf Not IsDate(msk_data_inicial.Text) Then
        MsgBox "Informe a data da inicial.", vbInformation, "Atenção!"
        msk_data_inicial.SetFocus
    ElseIf Not IsDate(msk_data_final.Text) Then
        MsgBox "Informe a data da final.", vbInformation, "Atenção!"
        msk_data_final.SetFocus
    ElseIf CDate(msk_data_final.Text) < CDate(msk_data_inicial.Text) Then
        MsgBox "A data final deve ser maior ou igual a " & msk_data_inicial.Text & ".", vbInformation, "Atenção!"
        msk_data_final.SetFocus
    ElseIf CDate(msk_data_baixa.Text) < CDate(msk_data_final.Text) Then
        MsgBox "A data da baixa deve ser maior ou igual a " & msk_data_final.Text & ".", vbInformation, "Atenção!"
        msk_data_baixa.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub dtcboCliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_baixa.SetFocus
    End If
End Sub
Private Sub dtcboCliente_LostFocus()
    If Cliente.LocalizarCodigo(Val(dtcboCliente.BoundText)) Then
        g_string = CalculaDataAbastecimentoVencimento(Cliente.CodigoVencimento, Date)
        If g_string <> "" Then
            If lClienteAnterior <> CLng(dtcboCliente.BoundText) Then
                lClienteAnterior = CLng(dtcboCliente.BoundText)
                msk_data_inicial.Text = RetiraGString(1)
                msk_data_final.Text = RetiraGString(2)
                msk_data_baixa.Text = Format(Date, "dd/mm/yyyy")
            End If
        End If
        g_string = ""
    Else
        MsgBox "Cliente não cadastrado!" & Chr(10) & Chr(10) & "Escolha outro cliente.", vbInformation, "Erro de Integridade!"
        dtcboCliente.SetFocus
    End If
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    If Len(g_string) > 0 Then
        If RetiraGString(1) = "ConfereNotaAbastecimento" Then
            dtcboCliente.BoundText = CLng(RetiraGString(3))
            msk_data_inicial.Text = RetiraGString(4)
            msk_data_final.Text = RetiraGString(4)
            msk_data_baixa.Text = RetiraGString(5)
            chkNotaConferida.Value = 1
            If Cliente.LocalizarCodigo(CLng(RetiraGString(3))) Then
            End If
            cmd_baixa_Click
            Exit Sub
        End If
    End If
    dtcboCliente.SetFocus
    If ConfiguracaoDiversa.LocalizarCodigo(1, "Baixar Notas Abast. Pelo Financeiro") Then
        If ConfiguracaoDiversa.Verdadeiro Then
            MsgBox "Esta operação somente será aceita pelo movimento financeiro.", vbInformation + vbOKOnly + vbExclamation, "Operação não Permitida!"
            cmd_sair_Click
        End If
    Else
        MsgBox "Esta operação somente será aceita pelo movimento financeiro.", vbInformation + vbOKOnly + vbExclamation, "Operação não Permitida!"
        cmd_sair_Click
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        cmd_sair_Click
    End If
End Sub
Private Sub Form_Load()
    Screen.MousePointer = 1
    CentraForm Me
    
    Set Me.adodcCliente.Recordset = Conectar.RsConexao("SELECT [Razao Social], Codigo FROM Cliente WHERE Inativo = " & preparaBooleano(False) & " ORDER BY [Razao Social]")
    
    'g_data = g_data_def
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_baixa_GotFocus()
    msk_data_baixa.SelStart = 0
    msk_data_baixa.SelLength = 5
End Sub
Private Sub msk_data_baixa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_baixa.SetFocus
    End If
End Sub
Private Sub msk_data_final_GotFocus()
    msk_data_final.SelStart = 0
    msk_data_final.SelLength = 5
End Sub
Private Sub msk_data_final_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data_baixa.SetFocus
    End If
End Sub
Private Sub msk_data_inicial_GotFocus()
    msk_data_inicial.SelStart = 0
    msk_data_inicial.SelLength = 5
End Sub
Private Sub msk_data_inicial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data_final.SetFocus
    End If
End Sub
