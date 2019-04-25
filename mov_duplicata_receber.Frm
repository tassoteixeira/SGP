VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form movimento_duplicata_receber 
   Caption         =   "Movimento das Duplicatas à Receber"
   ClientHeight    =   5745
   ClientLeft      =   885
   ClientTop       =   1680
   ClientWidth     =   7065
   Icon            =   "mov_duplicata_receber.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "mov_duplicata_receber.frx":030A
   ScaleHeight     =   5745
   ScaleWidth      =   7065
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   4020
      Picture         =   "mov_duplicata_receber.frx":0750
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Imprime esta duplicata."
      Top             =   4800
      Width           =   735
   End
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   120
      Picture         =   "mov_duplicata_receber.frx":1A2A
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Cria um novo registro."
      Top             =   4800
      Width           =   735
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   900
      Picture         =   "mov_duplicata_receber.frx":30BC
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Altera o registro atual."
      Top             =   4800
      Width           =   735
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   1680
      Picture         =   "mov_duplicata_receber.frx":45B6
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Exclui o registro atual."
      Top             =   4800
      Width           =   735
   End
   Begin VB.CommandButton cmd_pesquisa 
      Caption         =   "&Pesquisa"
      Height          =   855
      Left            =   2460
      Picture         =   "mov_duplicata_receber.frx":5C48
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Pesquisa um registro específico."
      Top             =   4800
      Width           =   735
   End
   Begin VB.CommandButton cmd_sair 
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3240
      Picture         =   "mov_duplicata_receber.frx":70BA
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   4800
      Width           =   735
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   5220
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frmDados 
      Enabled         =   0   'False
      Height          =   4635
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   6855
      Begin VB.ComboBox cboPortadorFinanceiro 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2610
         Width           =   6555
      End
      Begin VB.TextBox txt_valor_custo 
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   20
         Top             =   4260
         Width           =   1095
      End
      Begin VB.TextBox txt_valor_vencimento 
         Height          =   285
         Left            =   4920
         MaxLength       =   10
         TabIndex        =   18
         Top             =   3780
         Width           =   1095
      End
      Begin VB.TextBox txt_numero_duplicata 
         Height          =   285
         Left            =   120
         MaxLength       =   11
         TabIndex        =   14
         Top             =   3180
         Width           =   1095
      End
      Begin VB.TextBox txt_numero_documento 
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin MSMask.MaskEdBox msk_data_vencimento 
         Height          =   300
         Left            =   120
         TabIndex        =   16
         Top             =   3720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_data_periodo_final 
         Height          =   300
         Left            =   4920
         TabIndex        =   10
         Top             =   2040
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_data_periodo_inicial 
         Height          =   300
         Left            =   120
         TabIndex        =   8
         Top             =   2040
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_data_emissao 
         Height          =   300
         Left            =   120
         TabIndex        =   6
         Top             =   1500
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
         Left            =   2520
         Top             =   960
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
         Bindings        =   "mov_duplicata_receber.frx":874C
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   6555
         _ExtentX        =   11562
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Razao Social"
         BoundColumn     =   "Codigo"
         Text            =   "dtcboCliente"
      End
      Begin VB.Label Label3 
         Caption         =   "Portador &Financeiro"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   11
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Valor do Custo Bancário"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   19
         Top             =   4065
         Width           =   1875
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "&Valor do Vencimento"
         Height          =   195
         Index           =   1
         Left            =   4440
         TabIndex        =   17
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "D&ata do Vencimento"
         Height          =   195
         Index           =   14
         Left            =   120
         TabIndex        =   15
         Top             =   3525
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "&Número da Duplicata"
         Height          =   195
         Index           =   13
         Left            =   120
         TabIndex        =   13
         Top             =   2970
         Width           =   1635
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Data do Período &Final"
         Height          =   195
         Index           =   7
         Left            =   4320
         TabIndex        =   9
         Top             =   1860
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Data do Período &Inicial"
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   7
         Top             =   1840
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Número do Documento"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   160
         Width           =   1755
      End
      Begin VB.Label Label3 
         Caption         =   "C&liente"
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   555
      End
      Begin VB.Label Label3 
         Caption         =   "Data da &Emissão"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   5
         Top             =   1300
         Width           =   1335
      End
   End
   Begin VB.Frame frmAtualizacao 
      Caption         =   "Atualização de Duplicata"
      ForeColor       =   &H80000002&
      Height          =   1815
      Left            =   2280
      TabIndex        =   21
      Top             =   1740
      Visible         =   0   'False
      Width           =   2115
      Begin VB.TextBox txt_numero_duplicata_2 
         Height          =   285
         Left            =   120
         MaxLength       =   11
         TabIndex        =   25
         Top             =   1350
         Width           =   1095
      End
      Begin VB.TextBox txt_valor_custo_2 
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   23
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Número da &Duplicata"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   24
         Top             =   1140
         Width           =   1635
      End
      Begin VB.Label Label3 
         Caption         =   "&Valor do Custo Bancário"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   22
         Top             =   510
         Width           =   1875
      End
   End
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   4800
      TabIndex        =   34
      Top             =   4680
      Width           =   2175
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "mov_duplicata_receber.frx":8767
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "mov_duplicata_receber.frx":9C61
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "mov_duplicata_receber.frx":B15B
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "mov_duplicata_receber.frx":C5CD
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   5280
      Picture         =   "mov_duplicata_receber.frx":DB4F
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Confirma o registro atual."
      Top             =   4800
      Width           =   795
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   6180
      Picture         =   "mov_duplicata_receber.frx":F159
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Cancela o registro atual."
      Top             =   4800
      Width           =   795
   End
End
Attribute VB_Name = "movimento_duplicata_receber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Início de variáveis padrão para relatório
Dim lPagina As Integer
Dim lNomeArquivo As String
Dim lSQL As String
'Fim de variáveis padrão para relatório
Dim flag_movimento_duplicata_receber As Integer
Dim lOpcao As Integer
Dim lNumeroDocumento As Long
Dim l_empresa As Integer
Dim l_local As Integer
Dim l_nome_banco As String
Dim lDataEmissao As Date
Dim lNumeroMovimentoCaixa As Long

Private BaixaDuplicataReceber As New cBaixaDuplicataReceber
Private Cliente As New cCliente
Private Configuracao As New cConfiguracao
Private IntegracaoCaixa As New cIntegracaoCaixa
Private MovBaixaNotaAbastecimento As New cBaixaNotaAbastecimento
Private MovCaixa As New cMovimentoCaixa
Private MovDuplicataReceber As New cDuplicataReceber
Private MovNotaAbastecimento As New cMovimentoNotaAbastecimento
Private PortadorFinanceiro As New cPortadorFinanceiro
Private Sub AtivaBotoes()
    cmd_novo.Enabled = True
    cmd_alterar.Enabled = True
    If g_nivel_acesso <= 3 Then
        cmd_excluir.Enabled = True
    End If
    cmd_pesquisa.Enabled = True
    cmd_sair.Enabled = True
    cmd_imprimir.Enabled = True
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
    frm_move.Visible = True
End Sub
Private Sub Inclui()
    Dim xNumero As String
    lOpcao = 1
    DesativaBotoes
    cmd_novo.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    txt_numero_duplicata.Text = 0
    'busca último número da duplicata
    If MovDuplicataReceber.LocalizarUltimo(g_empresa) Then
        txt_numero_duplicata.Text = MovDuplicataReceber.NumeroDuplicata
        txt_numero_duplicata.Text = 0
    End If
    'busca último número da duplicata baixada se for maior
    If BaixaDuplicataReceber.LocalizarUltimo(g_empresa) Then
        xNumero = BaixaDuplicataReceber.UltimoNumeroDuplicata(g_empresa)
        If xNumero > txt_numero_duplicata.Text Then
            txt_numero_duplicata.Text = xNumero
            txt_numero_duplicata.Text = 0
        End If
    End If
    txt_numero_duplicata.Text = CDbl(txt_numero_duplicata.Text) + 1
    txt_numero_documento.Text = MovDuplicataReceber.ProximoNumeroDocumento(g_empresa)
End Sub
Function IncluiMovimentoCaixa() As Boolean
    Dim xComplemento As String
    IncluiMovimentoCaixa = False
    lNumeroMovimentoCaixa = 0
    
    If IntegracaoCaixa.LocalizarNome(g_empresa, "DUPLICATAS A RECEBER") Then
        xComplemento = Cliente.RazaoSocial
        MovCaixa.Empresa = g_empresa
        MovCaixa.Data = Format(msk_data_emissao.Text, "dd/mm/yyyy")
        MovCaixa.NumeroMovimento = 1
        MovCaixa.Valor = fValidaValor2(txt_valor_vencimento.Text)
        MovCaixa.NumeroDocumento = Val(txt_numero_documento.Text)
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
            MsgBox "Não foi integrado no caixa o valor=" & txt_valor_vencimento.Text, vbInformation, "Erro de Integridade"
        End If
    Else
        MsgBox "Não existe a integração=" & "DUPLICATAS A RECEBER" & ".", vbInformation, "Registro Inexistente"
    End If
End Function
Private Sub AtualTabe()
    lNumeroDocumento = Val(txt_numero_documento.Text)
    MovDuplicataReceber.Empresa = Val(g_empresa)
    MovDuplicataReceber.NumeroDocumento = Val(txt_numero_documento.Text)
    MovDuplicataReceber.CodigoCliente = Val(dtcboCliente.BoundText)
    MovDuplicataReceber.DataEmissao = Format(msk_data_emissao.Text, "dd/mm/yyyy")
    MovDuplicataReceber.DataPeriodoInicial = Format(msk_data_periodo_inicial.Text, "dd/mm/yyyy")
    MovDuplicataReceber.DataPeriodoFinal = Format(msk_data_periodo_final.Text, "dd/mm/yyyy")
    MovDuplicataReceber.NumeroDuplicata = txt_numero_duplicata.Text
    MovDuplicataReceber.NumeroContaBancaria = ""
    If cboPortadorFinanceiro.ItemData(cboPortadorFinanceiro.ListIndex) > 0 Then
        If PortadorFinanceiro.LocalizarCodigo(cboPortadorFinanceiro.ItemData(cboPortadorFinanceiro.ListIndex)) Then
            MovDuplicataReceber.NumeroContaBancaria = PortadorFinanceiro.NumeroContaBancaria
        End If
    End If
    MovDuplicataReceber.DataVencimento = Format(msk_data_vencimento.Text, "dd/mm/yyyy")
    MovDuplicataReceber.ValorVencimento = fValidaValor2(txt_valor_vencimento.Text)
    MovDuplicataReceber.ValorCustoBancario = fValidaValor2(txt_valor_custo.Text)
    MovDuplicataReceber.NumeroMovimentoCaixa = lNumeroMovimentoCaixa
End Sub
Private Sub AtualTela()
    Dim i As Integer
    
    lNumeroDocumento = MovDuplicataReceber.NumeroDocumento
    lDataEmissao = MovDuplicataReceber.DataEmissao
    lNumeroMovimentoCaixa = MovDuplicataReceber.NumeroMovimentoCaixa
    txt_numero_documento.Text = MovDuplicataReceber.NumeroDocumento
    dtcboCliente.BoundText = MovDuplicataReceber.CodigoCliente
    msk_data_emissao.Text = Format(MovDuplicataReceber.DataEmissao, "dd/mm/yyyy")
    msk_data_periodo_inicial.Text = Format(MovDuplicataReceber.DataPeriodoInicial, "dd/mm/yyyy")
    msk_data_periodo_final.Text = Format(MovDuplicataReceber.DataPeriodoFinal, "dd/mm/yyyy")
    cboPortadorFinanceiro.ListIndex = -1
    If MovDuplicataReceber.NumeroContaBancaria <> "" Then
        If PortadorFinanceiro.LocalizarNumeroContaBancaria(MovDuplicataReceber.NumeroContaBancaria) Then
            For i = 0 To cboPortadorFinanceiro.ListCount - 1
                cboPortadorFinanceiro.ListIndex = i
                'If cboPortadorFinanceiro.ItemData(i) = Bomba.NumeroIlha & Format(Bomba.Codigo, "00") Then
                If cboPortadorFinanceiro.Text = PortadorFinanceiro.Nome Then
                    Exit For
                End If
            Next
            'cboPortadorFinanceiro.ListIndex = PortadorFinanceiro.Codigo
        End If
    End If
    txt_numero_duplicata.Text = MovDuplicataReceber.NumeroDuplicata
    msk_data_vencimento.Text = Format(MovDuplicataReceber.DataVencimento, "dd/mm/yyyy")
    txt_valor_vencimento.Text = Format(MovDuplicataReceber.ValorVencimento, "###,###,##0.00")
    txt_valor_custo.Text = Format(MovDuplicataReceber.ValorCustoBancario, "###,###,##0.00")
    'txtNumeroContaBancaria.Text = MovDuplicataReceber.NumeroContaBancaria
    cmd_imprimir.Enabled = True
    If Cliente.LocalizarCodigo(MovDuplicataReceber.CodigoCliente) Then
        If Cliente.ImprimeDuplicata = False Then
            cmd_imprimir.Enabled = False
        End If
    End If
    frmDados.Enabled = False
End Sub
Private Sub DesativaBotoes()
    cmd_novo.Enabled = False
    cmd_alterar.Enabled = False
    cmd_excluir.Enabled = False
    cmd_pesquisa.Enabled = False
    cmd_sair.Enabled = False
    cmd_imprimir.Enabled = False
    frm_move.Visible = False
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
End Sub
Private Sub ExcluiMovimentoCaixa()
    If Not MovCaixa.Excluir(g_empresa, lDataEmissao, lNumeroMovimentoCaixa) Then
        MsgBox "Não foi excluído o movimento do caixa!", vbInformation, "Erro de Integridade."
    End If
End Sub
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set BaixaDuplicataReceber = Nothing
    Set Cliente = Nothing
    Set Configuracao = Nothing
    Set IntegracaoCaixa = Nothing
    Set MovBaixaNotaAbastecimento = Nothing
    Set MovCaixa = Nothing
    Set MovDuplicataReceber = Nothing
    Set MovNotaAbastecimento = Nothing
    Set PortadorFinanceiro = Nothing
End Sub
Private Sub cmd_alterar_Click()
    Call GravaAuditoria(1, Me.name, 3, "")
    lOpcao = 2
    DesativaBotoes
    cmd_alterar.Visible = True
    cmd_alterar.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    frmDados.Enabled = True
    txt_valor_vencimento.SetFocus
End Sub
Private Sub cmd_anterior_Click()
    Call GravaAuditoria(1, Me.name, 13, "")
    If MovDuplicataReceber.LocalizarAnterior Then
        AtualTela
    Else
        MsgBox "Início de Arquivo.", vbInformation, "Atenção!"
        cmd_proximo.SetFocus
    End If
End Sub
Private Sub cmd_cancelar_Click()
    Call GravaAuditoria(1, Me.name, 9, "")
    LimpaTela
    If MovDuplicataReceber.LocalizarCodigo(g_empresa, lNumeroDocumento) Then
        AtualTela
        AtivaBotoes
        cmd_alterar.SetFocus
    Else
        DesativaBotoes
        cmd_novo.Enabled = True
        cmd_sair.Enabled = True
        cmd_novo.SetFocus
    End If
    lOpcao = 0
End Sub
Private Sub LimpaTela()
    txt_numero_documento.Text = ""
    dtcboCliente.BoundText = ""
    msk_data_emissao.Text = "__/__/____"
    msk_data_periodo_inicial.Text = "__/__/____"
    msk_data_periodo_final.Text = "__/__/____"
    cboPortadorFinanceiro.ListIndex = -1
    txt_numero_duplicata.Text = ""
    msk_data_vencimento.Text = "__/__/____"
    txt_valor_vencimento.Text = ""
    txt_valor_custo.Text = "0,00"
End Sub
Private Sub cmd_excluir_Click()
    Call GravaAuditoria(1, Me.name, 4, "")
    If CLng(txt_numero_documento.Text) > 0 Then
        If (MsgBox("Deseja excluir este registro?", vbQuestion + vbYesNo + vbDefaultButton2, "Exclusão de Registro!")) = vbYes Then
            Call GravaAuditoria(1, Me.name, 10, "Vcto:" & msk_data_vencimento.Text & " Cliente:" & Val(dtcboCliente.BoundText) & " De:" & msk_data_periodo_inicial.Text & " a " & msk_data_periodo_final.Text & " Vlr:" & txt_valor_vencimento.Text)
            Call ExcluiMovimentoCaixa
            If Not MovDuplicataReceber.Excluir(g_empresa, CLng(txt_numero_documento.Text)) Then
                MsgBox "Não foi possível excluir este registro!", vbInformation, "Erro de Integridade."
            End If
            Call EstornaBaixaNotaAbastecimento(CLng(txt_numero_documento.Text), CDate(msk_data_periodo_inicial.Text), CDate(msk_data_periodo_final.Text))
            LimpaTela
            If MovDuplicataReceber.LocalizarUltimo(g_empresa) Then
                AtualTela
            Else
                DesativaBotoes
                cmd_novo.Enabled = True
                cmd_sair.Enabled = True
                cmd_novo.SetFocus
            End If
        End If
    End If
End Sub
Private Sub cmd_imprimir_Click()
    Dim xProximoNumero As Long
    Call GravaAuditoria(1, Me.name, 7, "Vcto:" & msk_data_vencimento.Text & " Cliente:" & Val(dtcboCliente.BoundText) & " De:" & msk_data_periodo_inicial.Text & " a " & msk_data_periodo_final.Text & " Vlr:" & txt_valor_vencimento.Text)
    If MovDuplicataReceber.NumeroDuplicata <> 1 Then
        xProximoNumero = MovDuplicataReceber.NumeroDuplicata
    Else
        xProximoNumero = MovDuplicataReceber.UltimoNumeroDuplicataConta(g_empresa, MovDuplicataReceber.NumeroContaBancaria) + 1
        MovDuplicataReceber.NumeroDuplicata = xProximoNumero
        If Not MovDuplicataReceber.Alterar(g_empresa, CLng(txt_numero_documento.Text)) Then
            MsgBox "Não foi possível alterar o número da duplicata bancária.", vbCritical, "Erro de Integridade!"
            Exit Sub
        End If
        txt_numero_duplicata.Text = MovDuplicataReceber.NumeroDuplicata
    End If
    gStringChamada = g_empresa & "|@|" & txt_numero_documento.Text & "|@|" & xProximoNumero & "|@|" & "IMPRIMIR" & "|@|"
    Call GravaSgpNetCadastroIni("GeraBoletoBancario")
    Exit Sub
    
    
    
    
    If SelecionaImpressoraEpsonDuplicata(Me) Then
        SelecionaPaginaDuplicata
        DesativaBotoes
        'passar para o final
        frmAtualizacao.Top = 1700
        frmAtualizacao.Left = 2280
        frmAtualizacao.Visible = True
        txt_numero_duplicata_2.Text = "1"
        If Configuracao.LocalizarCodigo(g_empresa) Then
            txt_valor_custo_2.Text = Format(Configuracao.CustoDuplicata, "###,##0.00")
            txt_numero_duplicata_2.Text = Configuracao.NumeroDuplicata + 1
        End If
        If MovDuplicataReceber.LocalizarCodigo(g_empresa, lNumeroDocumento) Then
            AtualTela
        Else
            MsgBox "Não foi possível localizar este registro!", vbInformation, "Erro de Integridade."
        End If
        txt_numero_duplicata_2.SetFocus
    End If
End Sub
Private Sub cmd_novo_Click()
    Call GravaAuditoria(1, Me.name, 2, "")
    LimpaTela
    Inclui
    frmDados.Enabled = True
    dtcboCliente.SetFocus
End Sub
Private Sub cmd_novo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 16 Then
        MsgBox "PROCESSAMENTO"
        Call ProcessaDuplicataReceber
    End If
End Sub
Private Sub cmd_ok_Click()
    On Error GoTo FileError
    If ValidaCampos Then
        If lOpcao = 1 Then
            Call GravaAuditoria(1, Me.name, 10, "Vcto:" & msk_data_vencimento.Text & " Cliente:" & Val(dtcboCliente.BoundText) & " De:" & msk_data_periodo_inicial.Text & " a " & msk_data_periodo_final.Text & " Vlr:" & txt_valor_vencimento.Text)
            If Not IncluiMovimentoCaixa Then
                MsgBox "Não foi possível integrar com o Caixa!", vbInformation, "Erro de Integridade."
            End If
            AtualTabe
            If MovDuplicataReceber.Incluir Then
                Call BaixaNotaAbastecimento(CDate(msk_data_periodo_inicial.Text), CDate(msk_data_periodo_final.Text))
            Else
                MsgBox "Não foi possível incluir este registro!", vbInformation, "Erro de Integridade."
                Exit Sub
            End If
        ElseIf lOpcao = 2 Then
            Call GravaAuditoria(1, Me.name, 10, "De: Vcto:" & Format(MovDuplicataReceber.DataVencimento, "dd/mm/yyyy") & " Cliente:" & MovDuplicataReceber.CodigoCliente & " De:" & Format(MovDuplicataReceber.DataPeriodoInicial, "dd/mm/yyyy") & " a " & Format(MovDuplicataReceber.DataPeriodoFinal, "dd/mm/yyyy") & " Vlr:" & MovDuplicataReceber.ValorVencimento)
            Call GravaAuditoria(1, Me.name, 10, "Para: Vcto:" & msk_data_vencimento.Text & " Cliente:" & Val(dtcboCliente.BoundText) & " De:" & msk_data_periodo_inicial.Text & " a " & msk_data_periodo_final.Text & " Vlr:" & txt_valor_vencimento.Text)
            Call ExcluiMovimentoCaixa
            If Not IncluiMovimentoCaixa Then
                MsgBox "Não foi possível integrar com o Caixa!", vbInformation, "Erro de Integridade."
            End If
            AtualTabe
            If Not MovDuplicataReceber.Alterar(g_empresa, lNumeroDocumento) Then
                MsgBox "Não foi possível alterar este registro!", vbInformation, "Erro de Integridade."
                Exit Sub
            End If
        End If
        AtivaBotoes
        lOpcao = 0
        If MovDuplicataReceber.LocalizarCodigo(g_empresa, lNumeroDocumento) Then
            AtualTela
        Else
            MsgBox "Não foi possível localizar este registro!", vbInformation, "Erro de Integridade."
        End If
        If cmd_imprimir.Enabled Then
            cmd_imprimir.SetFocus
        Else
            cmd_novo.SetFocus
        End If
    End If
    Exit Sub
FileError:
    'ErroArquivo tbl_duplicata_receber.Name, "Duplicata a Recebera"
    Exit Sub
End Sub
Function ValidaCampos() As Boolean
    ValidaCampos = False
    If CLng(txt_numero_documento) = 0 Then
        MsgBox "Informe o número do documento.", vbInformation, "Atenção!"
        txt_numero_documento.SetFocus
    ElseIf Val(dtcboCliente.BoundText) = 0 Then
        MsgBox "Escolha um cliente.", vbInformation, "Atenção!"
        dtcboCliente.SetFocus
    ElseIf Not IsDate(msk_data_emissao) Then
        MsgBox "Informe a data de emissao.", vbInformation, "Atenção!"
        msk_data_emissao.SetFocus
    ElseIf Not IsDate(msk_data_periodo_inicial) Then
        MsgBox "Informe a data do período inicial.", vbInformation, "Atenção!"
        msk_data_periodo_inicial.SetFocus
    ElseIf Not IsDate(msk_data_periodo_final) Then
        MsgBox "Informe a data do período final.", vbInformation, "Atenção!"
        msk_data_periodo_final.SetFocus
    ElseIf txt_numero_duplicata = "" Then
        MsgBox "Informe o número da duplicata.", vbInformation, "Atenção!"
        txt_numero_duplicata.SetFocus
    ElseIf Not IsDate(msk_data_vencimento) Then
        MsgBox "Informe a data do vencimento.", vbInformation, "Atenção!"
        msk_data_vencimento.SetFocus
    ElseIf fValidaValor2(txt_valor_vencimento) = 0 Then
        MsgBox "Informe o valor do vencimento.", vbInformation, "Atenção!"
        txt_valor_vencimento.SetFocus
    ElseIf CDate(msk_data_vencimento) < CDate(msk_data_emissao) Then
        MsgBox "A data do vencimento deve ser maior ou igual a " & msk_data_emissao & ".", vbInformation, "Atenção!"
        msk_data_vencimento.SetFocus
    ElseIf CDate(msk_data_emissao) < CDate(msk_data_periodo_final) Then
        MsgBox "A data de emissão deve ser maior ou igual a " & msk_data_periodo_final & ".", vbInformation, "Atenção!"
        msk_data_emissao.SetFocus
    ElseIf CDate(msk_data_periodo_final) < CDate(msk_data_periodo_inicial) Then
        MsgBox "A data de emissão deve ser maior ou igual a " & msk_data_periodo_inicial & ".", vbInformation, "Atenção!"
        msk_data_periodo_final.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_pesquisa_Click()
    Call GravaAuditoria(1, Me.name, 5, "")
    consulta_duplicata_receber.Show 1
    If Len(g_string) > 0 Then
        lNumeroDocumento = RetiraGString(1)
        If MovDuplicataReceber.LocalizarCodigo(g_empresa, lNumeroDocumento) Then
            AtualTela
        Else
            MsgBox "Não foi possível localizar este registro!", vbInformation, "Erro de Integridade."
        End If
    End If
End Sub
Private Sub cmd_primeiro_Click()
    Call GravaAuditoria(1, Me.name, 12, "")
    If MovDuplicataReceber.LocalizarPrimeiro Then
        AtualTela
        cmd_proximo.SetFocus
    Else
        MsgBox "Não há registros nesta empresa.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    Call GravaAuditoria(1, Me.name, 14, "")
    If MovDuplicataReceber.LocalizarProximo Then
        AtualTela
    Else
        MsgBox "Fim de Arquivo.", vbInformation, "Atenção!"
        cmd_anterior.SetFocus
    End If
End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub cmd_ultimo_Click()
    Call GravaAuditoria(1, Me.name, 15, "")
    If MovDuplicataReceber.LocalizarUltimo(g_empresa) Then
        AtualTela
        cmd_anterior.SetFocus
    Else
        MsgBox "Não há registros nesta empresa.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub dtcboCliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data_emissao.SetFocus
    End If
End Sub
Private Sub dtcboCliente_LostFocus()
    If (lOpcao = 1 Or lOpcao = 2) And Val(dtcboCliente.BoundText) > 0 Then
        If Cliente.LocalizarCodigo(Val(dtcboCliente.BoundText)) Then
            If Cliente.Inativo = True Then
                MsgBox "O cliente " & Trim(Cliente.RazaoSocial) & " está inativo.", vbInformation, "Atenção!"
                dtcboCliente.SetFocus
                Exit Sub
            End If
            If Cliente.Empresa <> g_empresa And Cliente.CodigoConvenio = 1 Then
                MsgBox "Este cliente é da empresa " & Cliente.Empresa & "." & Chr(10) & Chr(10) & "Mude o cliente escolhido.", vbInformation, "Erro de Integridade!"
                dtcboCliente.SetFocus
                Exit Sub
            Else
                g_string = CalculaDataAbastecimentoVencimento(Cliente.CodigoVencimento, Date)
                If g_string <> "" Then
                    msk_data_emissao.Text = Format(g_data_def, "dd/mm/yyyy")
                    msk_data_periodo_inicial.Text = RetiraGString(1)
                    msk_data_periodo_final.Text = RetiraGString(2)
                    msk_data_vencimento.Text = RetiraGString(3)
                End If
                g_string = ""
                txt_valor_vencimento_GotFocus
            End If
        Else
            MsgBox "Cliente não cadastrado!" & Chr(10) & Chr(10) & "Escolha outro cliente.", vbInformation, "Erro de Integridade!"
            dtcboCliente.SetFocus
        End If
    End If
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    If g_empresa <> l_empresa Then
        flag_movimento_duplicata_receber = 0
    End If
    If flag_movimento_duplicata_receber = 0 Then
        lOpcao = 0
        l_empresa = g_empresa
        DesativaBotoes
        If MovDuplicataReceber.LocalizarUltimo(g_empresa) Then
            AtivaBotoes
            AtualTela
        Else
            cmd_novo.Enabled = True
            cmd_sair.Enabled = True
        End If
        cmd_novo.SetFocus
    Else
        flag_movimento_duplicata_receber = 0
    End If
    Screen.MousePointer = 1
End Sub
Private Sub Form_Deactivate()
    flag_movimento_duplicata_receber = 1
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
    ElseIf KeyCode = vbKeyF5 And lOpcao = 0 Then
        KeyCode = 0
        cmd_pesquisa_Click
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
    CentraForm Me
    
    Set adodcCliente.Recordset = Conectar.RsConexao("SELECT Codigo, [Razao Social] FROM Cliente ORDER BY [Razao Social]")
    l_nome_banco = ReadINI("GRUPO DE EMPRESAS", "Nome do Banco para sair na Impressao de Duplicata", gArquivoIni)
    PreencheCboPortadorFinanceiro
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub

Private Sub msk_data_emissao_GotFocus()
    msk_data_emissao.SelStart = 0
    msk_data_emissao.SelLength = 5
End Sub
Private Sub msk_data_emissao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data_periodo_inicial.SetFocus
    End If
End Sub
Private Sub msk_data_periodo_final_GotFocus()
    msk_data_periodo_final.SelStart = 0
    msk_data_periodo_final.SelLength = 5
End Sub
Private Sub msk_data_periodo_final_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_numero_duplicata.SetFocus
    End If
End Sub
Private Sub msk_data_periodo_inicial_GotFocus()
    msk_data_periodo_inicial.SelStart = 0
    msk_data_periodo_inicial.SelLength = 5
End Sub
Private Sub msk_data_periodo_inicial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data_periodo_final.SetFocus
    End If
End Sub
Private Sub msk_data_vencimento_GotFocus()
    msk_data_vencimento.SelStart = 0
    msk_data_vencimento.SelLength = 5
End Sub
Private Sub msk_data_vencimento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_valor_vencimento.SetFocus
    End If
End Sub
Private Sub txt_numero_documento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dtcboCliente.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_numero_documento_LostFocus()
    If lOpcao = 1 And txt_numero_documento <> "" Then
        If MovDuplicataReceber.LocalizarCodigo(g_empresa, CLng(txt_numero_documento)) Then
            MsgBox "Já existe duplicata cadastrado com este número de documento." & Chr(10) & Chr(10) & "Mude o número do documento informado.", vbInformation, "Duplicidade de Registro!"
            txt_numero_documento.SetFocus
            Exit Sub
        End If
        If BaixaDuplicataReceber.LocalizarCodigo(g_empresa, CLng(txt_numero_documento.Text)) Then
            MsgBox "Existe duplicata baixada com este número de documento." & Chr(10) & Chr(10) & "Mude o número do documento informado.", vbInformation, "Duplicidade de Registro!"
            txt_numero_documento.SetFocus
            Exit Sub
        End If
    End If
End Sub
Private Sub txt_numero_duplicata_2_GotFocus()
    txt_numero_duplicata_2.SelStart = 0
    txt_numero_duplicata_2.SelLength = Len(txt_numero_duplicata_2.Text)
End Sub
Private Sub txt_numero_duplicata_2_KeyPress(KeyAscii As Integer)
Dim x_opcao As Integer
    If KeyAscii = 13 Then
        KeyAscii = 0
        AtivaBotoes
        x_opcao = (MsgBox("Deseja atualizar a alteração?", 3 + 32 + 0, "Atualização Automática!"))
        frmAtualizacao.Visible = False
        If MovDuplicataReceber.LocalizarCodigo(g_empresa, lNumeroDocumento) Then
            AtualTela
        Else
            MsgBox "Não foi possível localizar este registro!", vbInformation, "Erro de Integridade."
        End If
        If x_opcao = 6 Then
            MovDuplicataReceber.ValorCustoBancario = fValidaValor2(txt_valor_custo_2.Text)
            MovDuplicataReceber.NumeroDuplicata = txt_numero_duplicata_2.Text
            If Not MovDuplicataReceber.Alterar(g_empresa, lNumeroDocumento) Then
                MsgBox "Não foi possível alterar este registro!", vbInformation, "Erro de Integridade."
            End If
            Configuracao.NumeroDuplicata = CLng(txt_numero_duplicata_2.Text)
            If Not Configuracao.Alterar(g_empresa) Then
                MsgBox "Não foi possível alterar controle do número de duplicata!", vbInformation, "Erro de Integridade."
            End If
            If MovDuplicataReceber.LocalizarCodigo(g_empresa, lNumeroDocumento) Then
                AtualTela
            Else
                MsgBox "Não foi possível localizar este registro!", vbInformation, "Erro de Integridade."
            End If
        ElseIf x_opcao = 2 Then
            cmd_novo.SetFocus
            Exit Sub
        End If
        If UCase(g_nome_empresa) Like "*MARQUES*" Then
            RelatorioEsmeralda
        ElseIf UCase(g_nome_empresa) Like "*BOSQUE*" Then
            RelatorioBosque2
        Else
            Relatorio
        End If
        cmd_novo.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_numero_duplicata_2_LostFocus()
    If txt_numero_duplicata_2.Text <> "" Then
        If CDbl(txt_numero_duplicata_2.Text) > 1 Then
            If MovDuplicataReceber.LocalizarNumeroDuplicata(g_empresa, txt_numero_duplicata_2.Text) Then
                If CLng(MovDuplicataReceber.NumeroDocumento) <> CLng(txt_numero_documento.Text) Then
                    If MovDuplicataReceber.Empresa = g_empresa And CDbl(MovDuplicataReceber.NumeroDuplicata) = CDbl(txt_numero_duplicata_2.Text) Then
                        MsgBox "Já existe duplicata cadastrado com este número de duplicata." & Chr(10) & Chr(10) & "Mude o número da duplicata.", vbInformation, "Duplicidade de Registro!"
                        txt_numero_duplicata_2.SetFocus
                    End If
                End If
            End If
        End If
        
        
        If CDbl(txt_numero_duplicata_2.Text) > 1 Then
            If BaixaDuplicataReceber.LocalizarNumeroDuplicata(g_empresa, CDbl(txt_numero_duplicata_2.Text)) Then
                MsgBox "Existe duplicata baixada com este número de duplicata." & Chr(10) & Chr(10) & "Mude o número da duplicata.", vbInformation, "Duplicidade de Registro!"
                txt_numero_duplicata_2.SetFocus
            End If
        End If
    End If
End Sub
Private Sub txt_numero_duplicata_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data_vencimento.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_numero_duplicata_LostFocus()
    If lOpcao = 1 And txt_numero_duplicata.Text <> "" Then
        If CDbl(txt_numero_duplicata.Text) > 1 Then
            If MovDuplicataReceber.LocalizarNumeroDuplicata(g_empresa, txt_numero_duplicata.Text) Then
                MsgBox "Já existe duplicata cadastrado com este número de duplicata." & Chr(10) & Chr(10) & "Mude o número da duplicata.", vbInformation, "Duplicidade de Registro!"
                txt_numero_duplicata.SetFocus
            End If
        End If
        If CDbl(txt_numero_duplicata.Text) > 1 Then
            If BaixaDuplicataReceber.LocalizarNumeroDuplicata(g_empresa, txt_numero_duplicata.Text) Then
                MsgBox "Existe duplicata baixada com este número de duplicata." & Chr(10) & Chr(10) & "Mude o número da duplicata.", vbInformation, "Duplicidade de Registro!"
                txt_numero_duplicata.SetFocus
            End If
        End If
    End If
End Sub
Private Sub txt_valor_custo_2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{Tab}"
    End If
End Sub
Private Sub txt_valor_custo_2_LostFocus()
    txt_valor_custo_2.Text = Format(txt_valor_custo_2.Text, "###,###,##0.00")
End Sub
Private Sub txt_valor_custo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub txt_valor_custo_LostFocus()
    txt_valor_custo.Text = Format(txt_valor_custo.Text, "###,###,##0.00")
End Sub
Private Sub txt_valor_vencimento_GotFocus()
    If lOpcao = 1 And dtcboCliente.BoundText <> "" Then
        'txt_valor_vencimento.Text = Format(MovNotaAbastecimento.TotalDataLiquido(g_empresa, CLng(dtcboCliente.BoundText), CDate(msk_data_periodo_inicial.Text), CDate(msk_data_periodo_final.Text), False), "###,###,##0.00")
        txt_valor_vencimento.Text = Format(MovNotaAbastecimento.TotalDataLiquido(0, CLng(dtcboCliente.BoundText), CDate(msk_data_periodo_inicial.Text), CDate(msk_data_periodo_final.Text), False), "###,###,##0.00")
    End If
    txt_valor_vencimento.SelStart = 0
    txt_valor_vencimento.SelLength = Len(txt_valor_vencimento.Text)
End Sub
Private Sub txt_valor_vencimento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_valor_custo.SetFocus
    End If
End Sub
Private Sub txt_valor_vencimento_LostFocus()
    txt_valor_vencimento = Format(txt_valor_vencimento, "###,###,##0.00")
End Sub
Private Sub PreencheCboPortadorFinanceiro()
    Dim rsPortadorFinanceiro As adodb.Recordset
    Dim xSQL As String
    
    cboPortadorFinanceiro.Clear
    cboPortadorFinanceiro.AddItem " Sem Portador Financeiro"
    cboPortadorFinanceiro.ItemData(cboPortadorFinanceiro.NewIndex) = 0
    
    Set rsPortadorFinanceiro = New adodb.Recordset
    xSQL = "SELECT Codigo, Nome"
    xSQL = xSQL & "  FROM PortadorFinanceiro"
    xSQL = xSQL & " WHERE Empresa = " & g_empresa
    xSQL = xSQL & "   AND [Instituicao Financeira] = " & preparaBooleano(True)
    xSQL = xSQL & " ORDER BY Nome, Codigo"
    Set rsPortadorFinanceiro = Conectar.RsConexao(xSQL)
    If rsPortadorFinanceiro.RecordCount > 0 Then
        Do Until rsPortadorFinanceiro.EOF
            cboPortadorFinanceiro.AddItem rsPortadorFinanceiro("Nome").Value
            cboPortadorFinanceiro.ItemData(cboPortadorFinanceiro.NewIndex) = rsPortadorFinanceiro("Codigo").Value
            rsPortadorFinanceiro.MoveNext
        Loop
    End If
    rsPortadorFinanceiro.Close
    Set rsPortadorFinanceiro = Nothing
End Sub
Private Sub ProcessaDuplicataReceber()
    Dim xData As Date
    Dim rsDuplicataReceber As adodb.Recordset
    Dim xSQL As String
    
    On Error GoTo FileError
    
    xData = CDate("01/10/2004")
    Set rsDuplicataReceber = New adodb.Recordset
    xSQL = "SELECT [Numero do Documento], [Data de Emissao] FROM Duplicata_Receber WHERE Empresa = " & g_empresa & " AND [Data de Emissao] >= " & xData & " ORDER BY [Data de Emissao]"
    Set rsDuplicataReceber = Conectar.RsConexao(xSQL)
    With rsDuplicataReceber
        If .RecordCount > 0 Then
            Do Until .EOF
                If MovDuplicataReceber.LocalizarCodigo(g_empresa, rsDuplicataReceber("Numero do Documento").Value) Then
                    AtualTela
                    If Not IncluiMovimentoCaixa Then
                        MsgBox "Não foi possível incluir este registro no Caixa!", vbInformation, "Erro de Integridade."
                    Else
                        MovDuplicataReceber.NumeroMovimentoCaixa = lNumeroMovimentoCaixa
                        If Not MovDuplicataReceber.Alterar(g_empresa, MovDuplicataReceber.NumeroDocumento) Then
                            MsgBox "Não foi possível alterar este registro!", vbInformation, "Erro de Integridade."
                        End If
                    End If
                End If
                .MoveNext
            Loop
        Else
            MsgBox "Não há registros nesta empresa.", vbInformation, "Erro de Verificação!"
        End If
    End With
    rsDuplicataReceber.Close
    Set rsDuplicataReceber = Nothing
    MsgBox "Processamento Concluído!"
    Exit Sub
FileError:
    MsgBox "Erro ao processar Duplicata Receber", vbInformation, "ProcessaDuplicataReceber"
End Sub



Private Sub Relatorio()
    Dim posicao_y As Currency
    Dim tamanho_form As Integer
    Dim largura_form As Integer
    Dim x_linha As Currency
    Dim x_tamanho_string As Currency
    Dim x_extenso As String
    Dim x_asterisco As String
    Dim xJuros As Currency
    l_local = 1
    x_linha = 0.4
    'seleciona medidas para polegadas
    Printer.ScaleMode = 5
    'Seleciona largura do formulário
    Printer.ScaleWidth = 9
    largura_form = Printer.ScaleWidth
    'Seleciona altura do formulário
    Printer.ScaleHeight = 4
    tamanho_form = Printer.ScaleHeight
    'seleciona medidas para centímetros
    Printer.ScaleMode = 7
    'Seleciona nome da fonte
    Printer.FontName = "Arial"
    
    Printer.FontName = "Roman 10cpi"
'    ImprimeCentralizado Format(tbl_contas_pagar!data_vencimento, "dd/mm/yyyy"), l_coluna(3), l_coluna(4), l_linha_tab, l_local
    
'    ImprimeTexto "Pagável em qualquer banco até o vencimento", 0, 12, x_linha + 0.2, l_local
'    ImprimeTexto "Pagável em qualquer banco até o vencimento", 0, 12, x_linha + 0.215, l_local
    Printer.FontName = "Arial"
    Printer.FontBold = True
    Printer.FontSize = 12
    ImprimeValor msk_data_vencimento, 13, 15.6, x_linha, l_local
    
    Printer.FontName = "Roman 10cpi"
    
    ImprimeTexto "Pagável em qualquer banco até o vencimento", 0, 10, x_linha, l_local
    ImprimeTexto "Pagável em qualquer banco até o vencimento", 0, 10, x_linha, l_local
    ImprimeTexto Format(msk_data_emissao, "dd/mm/yy"), 0, 3, x_linha + 1.35, l_local
    ImprimeTexto Format(msk_data_emissao, "dd/mm/yy"), 0, 3, x_linha + 1.365, l_local
    ImprimeCentralizado Format(txt_numero_documento, "##,###,##0"), 1.8, 6.5, x_linha + 1.35, l_local
    ImprimeCentralizado Format(txt_numero_documento, "##,###,##0"), 1.8, 6.5, x_linha + 1.365, l_local
    ImprimeCentralizado "DUP.", 6.7, 8.6, x_linha + 1.35, l_local
    ImprimeCentralizado "DUP.", 6.7, 8.6, x_linha + 1.365, l_local
    ImprimeCentralizado "N", 8.5, 9.6, x_linha + 1.35, l_local
    ImprimeCentralizado "N", 8.5, 9.6, x_linha + 1.365, l_local
    
    Printer.FontName = "Arial"
    Printer.FontBold = True
    Printer.FontSize = 12
    ImprimeValor Format(fValidaValor2(txt_valor_vencimento.Text), "##,###,##0.00"), 12, 15.6, x_linha + 2.1, l_local
    
    ImprimeValor Trim(txt_valor_custo_2.Text), 14, 15.6, x_linha + 4.5, l_local
    
    'xJuros = fValidaValor2(txt_valor_custo_2.Text) + fValidaValor2(txt_valor_vencimento.Text)
    'ImprimeValor Format(xJuros, "##,###,##0.00"), 14, 15.6, x_linha + 5.1, l_local

    Printer.FontBold = False
    Printer.FontName = "Roman 10cpi"
    Printer.FontSize = 9
    ImprimeTexto "NOTAS REFERENTE AO PERÍODO DE.: " & msk_data_periodo_inicial & " A " & msk_data_periodo_final, 0, 10, x_linha + 3.3, l_local
    ImprimeTexto "NOTAS REFERENTE AO PERÍODO DE.: " & msk_data_periodo_inicial & " A " & msk_data_periodo_final, 0, 10, x_linha + 3.315, l_local
    
    ImprimeTexto "APÓS O VENCIMENTO: PAGÁVEL APENAS NO " & l_nome_banco, 0, 10, x_linha + 3.7, l_local
    ImprimeTexto "APÓS O VENCIMENTO: PAGÁVEL APENAS NO " & l_nome_banco, 0, 10, x_linha + 3.715, l_local
    
    xJuros = Format(fValidaValor2(txt_valor_vencimento.Text) * 0.5 / 100, "00000000.00")
    ImprimeTexto "COBRAR R$ " & Format(xJuros, "##,###,##0.00") & " POR DIA DE ATRASO", 0, 10, x_linha + 4.1, l_local
    ImprimeTexto "COBRAR R$ " & Format(xJuros, "##,###,##0.00") & " POR DIA DE ATRASO", 0, 10, x_linha + 4.1093, l_local
    
    ImprimeTexto "PROTESTAR APÓS 5 DIAS DE VENCIMENTO", 0, 10, x_linha + 4.8, l_local
    ImprimeTexto "PROTESTAR APÓS 5 DIAS DE VENCIMENTO", 0, 10, x_linha + 4.813, l_local
    
    ImprimeTexto "TX.BANC.", 11.7, 14, x_linha + 4.5, l_local
    
    
    ImprimeValor Format(dtcboCliente.BoundText, "##,###"), 0, 1, x_linha + 5.85, l_local
    ImprimeValor Format(dtcboCliente.BoundText, "##,###"), 0, 1, x_linha + 5.86, l_local
    ImprimeTexto dtcboCliente.Text, 1.1, 10, x_linha + 5.85, l_local
    ImprimeTexto dtcboCliente.Text, 1.1, 10, x_linha + 5.86, l_local
    
    If Cliente.LocalizarCodigo(Val(dtcboCliente.BoundText)) Then
        ImprimeTexto Trim(Cliente.Endereco) & " " & Cliente.Bairro, 1.1, 10.1, x_linha + 6.25, l_local
        ImprimeTexto Trim(Cliente.Endereco) & " " & Cliente.Bairro, 1.1, 10.1, x_linha + 6.265, l_local
        ImprimeTexto Mid(Cliente.CEP, 1, 2) & "." & Mid(Cliente.CEP, 3, 3) & "-" & Mid(Cliente.CEP, 6, 3), 10.2, 11.9, x_linha + 6.25, l_local
        ImprimeTexto Mid(Cliente.CEP, 1, 2) & "." & Mid(Cliente.CEP, 3, 3) & "-" & Mid(Cliente.CEP, 6, 3), 10.2, 11.9, x_linha + 6.265, l_local
        ImprimeTexto " - " & Cliente.Cidade & " " & Cliente.UF, 11.9, 17, x_linha + 6.25, l_local
        ImprimeTexto " - " & Cliente.Cidade & " " & Cliente.UF, 11.9, 17, x_linha + 6.265, l_local
        If Trim(Cliente.CGC) <> "" Then
            ImprimeTexto "CNPJ: " & fMascaraCNPJ(Cliente.CGC), 0, 5, x_linha + 6.65, l_local
            ImprimeTexto "I.E.: " & Cliente.InscricaoEstadual, 5, 10, x_linha + 6.65, l_local
        Else
            ImprimeTexto "CPF.: " & Cliente.CPF, 0, 5, x_linha + 6.65, l_local
            ImprimeTexto "R.G.: " & Cliente.Identidade, 5, 10, x_linha + 6.65, l_local
        End If
        ImprimeTexto "Fone: " & Cliente.Telefone, 10, 15, x_linha + 6.65, l_local
    End If
    Printer.EndDoc
End Sub



Private Sub RelatorioBosque()
    Dim posicao_y As Currency
    Dim tamanho_form As Integer
    Dim largura_form As Integer
    Dim x_linha As Currency
    Dim x_tamanho_string As Currency
    Dim x_extenso As String
    Dim x_asterisco As String
    Dim xJuros As Currency
    
    l_local = 1
    x_linha = 0.4
    'seleciona medidas para polegadas
    Printer.ScaleMode = 5
    'Seleciona largura do formulário
    Printer.ScaleWidth = 9
    largura_form = Printer.ScaleWidth
    'Seleciona altura do formulário
    Printer.ScaleHeight = 4
    tamanho_form = Printer.ScaleHeight
    'seleciona medidas para centímetros
    Printer.ScaleMode = 7
    'Seleciona nome da fonte
    Printer.FontName = "Arial"
    
    Printer.FontName = "Roman 10cpi"
'    ImprimeCentralizado Format(tbl_contas_pagar!data_vencimento, "dd/mm/yyyy"), l_coluna(3), l_coluna(4), l_linha_tab, l_local
    
'    ImprimeTexto "Pagável em qualquer banco até o vencimento", 0, 12, x_linha + 0.2, l_local
'    ImprimeTexto "Pagável em qualquer banco até o vencimento", 0, 12, x_linha + 0.215, l_local
    Printer.FontName = "Arial"
    Printer.FontBold = True
    Printer.FontSize = 12
    ImprimeValor msk_data_vencimento.Text, 13, 17.9, x_linha + 0.6, l_local
    
    Printer.FontName = "Roman 10cpi"
    
    'ImprimeTexto "Pagável em qualquer banco até o vencimento", 0, 10, x_linha, l_local
    'ImprimeTexto "Pagável em qualquer banco até o vencimento", 0, 10, x_linha, l_local
    ImprimeTexto Format(msk_data_emissao.Text, "dd/mm/yy"), 0, 3, x_linha + 2.15, l_local
    'ImprimeTexto Format(msk_data_emissao.Text, "dd/mm/yy"), 0, 3, x_linha + 1.965, l_local
    'ImprimeCentralizado Format(txt_numero_documento.Text, "##,###,##0"), 1.8, 6.5, x_linha + 1.35, l_local
    'ImprimeCentralizado Format(txt_numero_documento.Text, "##,###,##0"), 1.8, 6.5, x_linha + 1.365, l_local
    'ImprimeCentralizado "DUP.", 6.7, 8.6, x_linha + 1.35, l_local
    'ImprimeCentralizado "DUP.", 6.7, 8.6, x_linha + 1.365, l_local
    'ImprimeCentralizado "N", 8.5, 9.6, x_linha + 1.35, l_local
    'ImprimeCentralizado "N", 8.5, 9.6, x_linha + 1.365, l_local
    ImprimeTexto Format(msk_data_emissao.Text, "dd/mm/yy"), 10.5, 13, x_linha + 2.15, l_local
    'ImprimeTexto Format(msk_data_emissao.Text, "dd/mm/yy"), 10.5, 13, x_linha + 1.965, l_local
    
    Printer.FontName = "Arial"
    Printer.FontBold = True
    Printer.FontSize = 12
    ImprimeValor Format(fValidaValor2(txt_valor_vencimento.Text), "##,###,##0.00"), 13, 17.9, x_linha + 2.7, l_local
    
    'ImprimeValor Trim(txt_valor_custo_2.Text), 14, 15.6, x_linha + 4.5, l_local
    
    'xJuros = fValidaValor2(txt_valor_custo_2.Text) + fValidaValor2(txt_valor_vencimento.Text)
    'ImprimeValor Format(xJuros, "##,###,##0.00"), 14, 15.6, x_linha + 5.1, l_local

    Printer.FontBold = False
    Printer.FontName = "Arial"
    Printer.FontSize = 10
    ImprimeTexto "NOTAS REFERENTE AO PERÍODO DE.: " & msk_data_periodo_inicial.Text & " A " & msk_data_periodo_final.Text, 0, 14, x_linha + 3.6, l_local
    'ImprimeTexto "NOTAS REFERENTE AO PERÍODO DE.: " & msk_data_periodo_inicial.Text & " A " & msk_data_periodo_final.Text, 0, 14, x_linha + 3.615, l_local
    
    'ImprimeTexto "APÓS O VENCIMENTO: PAGÁVEL APENAS NO " & l_nome_banco, 0, 10, x_linha + 3.7, l_local
    'ImprimeTexto "APÓS O VENCIMENTO: PAGÁVEL APENAS NO " & l_nome_banco, 0, 10, x_linha + 3.715, l_local
    
    xJuros = Format(fValidaValor2(txt_valor_vencimento.Text) * 0.5 / 100, "00000000.00")
    ImprimeTexto "JUROS DE MORA 0,26% AO DIA: MULTA DE 3%", 0, 10, x_linha + 4.3, l_local
    'ImprimeTexto "COBRAR R$ " & Format(xJuros, "##,###,##0.00") & " POR DIA DE ATRASO", 0, 10, x_linha + 4.3093, l_local
    
    ImprimeTexto "PROTESTAR APÓS 10 DIAS DO VENCIMENTO", 0, 10, x_linha + 5, l_local
    'ImprimeTexto "PROTESTAR APÓS 10 DIAS DO VENCIMENTO", 0, 10, x_linha + 5.013, l_local
    
    'ImprimeTexto "TX.BANC.", 11.7, 14, x_linha + 4.5, l_local
    
    
    ImprimeValor Format(dtcboCliente.BoundText, "##,###"), 1.1, 2.1, x_linha + 6.25, l_local
    'ImprimeValor Format(dtcboCliente.BoundText, "##,###"), 1.1, 2.1, x_linha + 6.265, l_local
    ImprimeTexto dtcboCliente.Text, 2.2, 13, x_linha + 6.25, l_local
    'ImprimeTexto dtcboCliente.Text, 2.2, 13, x_linha + 6.265, l_local
    
    If Cliente.LocalizarCodigo(Val(dtcboCliente.BoundText)) Then
        ImprimeTexto Trim(Cliente.Endereco) & " " & Cliente.Bairro, 1.1, 10.1, x_linha + 6.65, l_local
        'ImprimeTexto Trim(Cliente.Endereco) & " " & Cliente.Bairro, 1.1, 10.1, x_linha + 6.665, l_local
        ImprimeTexto Mid(Cliente.CEP, 1, 2) & "." & Mid(Cliente.CEP, 3, 3) & "-" & Mid(Cliente.CEP, 6, 3), 10.2, 11.9, x_linha + 6.65, l_local
        'ImprimeTexto Mid(Cliente.CEP, 1, 2) & "." & Mid(Cliente.CEP, 3, 3) & "-" & Mid(Cliente.CEP, 6, 3), 10.2, 11.9, x_linha + 6.665, l_local
        ImprimeTexto " - " & Cliente.Cidade & " " & Cliente.UF, 11.9, 17, x_linha + 6.65, l_local
        'ImprimeTexto " - " & Cliente.Cidade & " " & Cliente.UF, 11.9, 17, x_linha + 6.665, l_local
        If Trim(Cliente.CGC) <> "" Then
            ImprimeTexto "CNPJ: " & fMascaraCNPJ(Cliente.CGC), 1.1, 6, x_linha + 7.05, l_local
            ImprimeTexto "I.E.: " & Cliente.InscricaoEstadual, 6, 11, x_linha + 7.05, l_local
        Else
            ImprimeTexto "CPF.: " & Cliente.CPF, 1.1, 5, x_linha + 7.05, l_local
            ImprimeTexto "R.G.: " & Cliente.Identidade, 6, 11, x_linha + 7.05, l_local
        End If
        ImprimeTexto "Fone: " & Cliente.Telefone, 11, 15, x_linha + 7.05, l_local
    End If
    Printer.EndDoc
End Sub



Private Sub RelatorioBosque2()
    'Dim posicao_y As Currency
    'Dim tamanho_form As Integer
    'Dim largura_form As Integer
    'Dim x_linha As Currency
    'Dim x_tamanho_string As Currency
    Dim x_extenso As String
    Dim x_asterisco As String
    'Dim xJuros As Currency
    
    Dim xLinha As String
    Dim i As Integer
    
    Open "lpt1" For Output As #1
    
    Print #1, Chr(27) & "0"        ' Muda o passo p/ 8 LPP
    Print #1, Chr(27) & Chr(15)    ' Condensado
    Print #1, Chr(27) & Chr(69)    ' Negrito
    
    
    '         12345678901234567890123456789012345678901234567890123456789012345678901234567890
    xLinha = "                                                                                "
    Mid(xLinha, 63, 10) = Format(msk_data_vencimento.Text, "dd/mm/yy")
    Print #1, Chr(27) & "E" & xLinha
    
    
    xLinha = "                                                                                "
    Print #1, xLinha
    Print #1, xLinha
    Print #1, xLinha
    xLinha = "                                                                                "
    Mid(xLinha, 1, 10) = Format(msk_data_emissao.Text, "dd/mm/yy")
    Mid(xLinha, 42, 10) = Format(msk_data_emissao.Text, "dd/mm/yy")
    Print #1, xLinha
    '         12345678901234567890123456789012345678901234567890123456789012345678901234567890
    xLinha = "                                                                                "
    Print #1, xLinha
    i = Len(Format(fValidaValor2(txt_valor_vencimento.Text), "##,###,##0.00"))
    Mid(xLinha, 57 + 13 - i, i) = Format(fValidaValor2(txt_valor_vencimento.Text), "##,###,##0.00")
    Print #1, xLinha
    
    
    xLinha = "                                                                                "
    Print #1, xLinha
    Print #1, xLinha
    'xLinha = "NOTAS REFERENTE AO PERIODO DE.: " & Format(msk_data_periodo_inicial.Text, "dd/mm/yy") & " A " & Format(msk_data_periodo_final.Text, "dd/mm/yy")
    Print #1, xLinha

    
    xLinha = "                                                                                "
    xLinha = "JUROS DE MORA 0,26% AO DIA: MULTA DE 3%"
    Print #1, xLinha
    
    xLinha = "                                                                                "
    xLinha = "PROTESTAR APOS 10 DIAS DO VENCIMENTO"
    Print #1, xLinha
    
    '         12345678901234567890123456789012345678901234567890123456789012345678901234567890
    xLinha = "                                                                                "
    Print #1, xLinha
    Print #1, xLinha
    Print #1, xLinha
    Print #1, xLinha
    Print #1, xLinha
    Print #1, Chr(27) & Chr(15)
    i = Len(Format(dtcboCliente.BoundText, "##,###"))
    Mid(xLinha, 1 + i, i) = Format(dtcboCliente.BoundText, "##,###")
    Mid(xLinha, 8, 40) = dtcboCliente.Text
    Print #1, xLinha
    
    If Cliente.LocalizarCodigo(Val(dtcboCliente.BoundText)) Then
        '         12345678901234567890123456789012345678901234567890123456789012345678901234567890
        xLinha = "                                                                                "
        Mid(xLinha, 4, 76) = Trim(Cliente.Endereco) & " " & Cliente.Bairro & " - " & Cliente.Cidade & " " & Cliente.UF
        Print #1, xLinha
        'xLinha = "                                                                                "
        'Mid(xLinha, 20, 20) = Mid(Cliente.CEP, 1, 2) & "." & Mid(Cliente.CEP, 3, 3) & "-" & Mid(Cliente.CEP, 6, 3)
        If Trim(Cliente.CGC) <> "" Then
            '         12345678901234567890123456789012345678901234567890123456789012345678901234567890
            xLinha = "                                                                                "
            Mid(xLinha, 1, 72) = "CNPJ: " & fMascaraCNPJ(Cliente.CGC) & "  I.E.: " & Cliente.InscricaoEstadual
        Else
            '         12345678901234567890123456789012345678901234567890123456789012345678901234567890
            xLinha = "                                                                                "
            Mid(xLinha, 1, 72) = "CPF.: " & Cliente.CPF & "   R.G.: " & Cliente.Identidade
        End If
        Mid(xLinha, 50, 20) = "Fone: " & fMascaraTelefone(Cliente.Telefone)
        Print #1, xLinha
    End If

    'Print #1, Chr(27) & Chr(8)     ' Tira Condensado
    'Print #1, Chr(27) & Chr(69)    ' Tira Negrito
    xLinha = "                                                                                "
    Print #1, xLinha
    Print #1, xLinha
    Print #1, xLinha
    Print #1, xLinha
    Print #1, xLinha
    Print #1, xLinha
    Print #1, xLinha
    Print #1, xLinha
    Close #1
End Sub
Private Sub RelatorioEsmeralda()
    Dim posicao_y As Currency
    Dim tamanho_form As Integer
    Dim largura_form As Integer
    Dim x_linha As Currency
    Dim x_tamanho_string As Currency
    Dim x_extenso As String
    Dim x_asterisco As String
    Dim xJuros As Currency
    l_local = 1
    x_linha = 0.4
    'seleciona medidas para polegadas
    Printer.ScaleMode = 5
    'Seleciona largura do formulário
    Printer.ScaleWidth = 9
    largura_form = Printer.ScaleWidth
    'Seleciona altura do formulário
    Printer.ScaleHeight = 4
    tamanho_form = Printer.ScaleHeight
    'seleciona medidas para centímetros
    Printer.ScaleMode = 7
    'Seleciona nome da fonte
    Printer.FontName = "Arial"
    
    '''Printer.FontName = "Roman 10cpi"
    Printer.FontSize = 10
'    ImprimeCentralizado Format(tbl_contas_pagar!data_vencimento, "dd/mm/yyyy"), l_coluna(3), l_coluna(4), l_linha_tab, l_local
    
'    ImprimeTexto "Pagável em qualquer banco até o vencimento", 0, 12, x_linha + 0.2, l_local
'    ImprimeTexto "Pagável em qualquer banco até o vencimento", 0, 12, x_linha + 0.215, l_local
    Printer.FontName = "Arial"
    Printer.FontBold = True
    Printer.FontSize = 12
    ImprimeValor msk_data_vencimento, 13, 20.1, x_linha - 0.3, l_local
    
    '''Printer.FontName = "Roman 10cpi"
    Printer.FontName = "Arial"
    Printer.FontBold = False
    Printer.FontSize = 10
    
    'ImprimeTexto "Pagável em qualquer banco até o vencimento", 0, 10, x_linha, l_local
    'ImprimeTexto "Pagável em qualquer banco até o vencimento", 0, 10, x_linha, l_local
    ImprimeTexto Format(msk_data_emissao, "dd/mm/yy"), 1, 4, x_linha + 1.25, l_local
    ImprimeTexto Format(msk_data_emissao, "dd/mm/yy"), 1, 4, x_linha + 1.265, l_local
    ImprimeCentralizado Format(txt_numero_documento, "##,###,##0"), 3.8, 9.5, x_linha + 1.25, l_local
    ImprimeCentralizado Format(txt_numero_documento, "##,###,##0"), 6.8, 9.5, x_linha + 1.265, l_local
    ImprimeCentralizado "DUP.", 9.2, 10.5, x_linha + 1.25, l_local
    ImprimeCentralizado "DUP.", 9.2, 10.5, x_linha + 1.265, l_local
    ImprimeCentralizado "N", 11.5, 12.6, x_linha + 1.25, l_local
    ImprimeCentralizado "N", 11.5, 12.6, x_linha + 1.265, l_local
    
    Printer.FontName = "Arial"
    Printer.FontBold = True
    Printer.FontSize = 12
    ImprimeValor Format(fValidaValor2(txt_valor_vencimento.Text), "##,###,##0.00"), 12, 20.1, x_linha + 1.6, l_local
    
    'ImprimeValor Trim(txt_valor_custo_2.Text), 14, 15.6, x_linha + 4.5, l_local
    
    'xJuros = fValidaValor2(txt_valor_custo_2.Text) + fValidaValor2(txt_valor_vencimento.Text)
    'ImprimeValor Format(xJuros, "##,###,##0.00"), 14, 15.6, x_linha + 5.1, l_local

    Printer.FontBold = False
    Printer.FontName = "Arial"
    Printer.FontSize = 9
    ImprimeTexto "NOTAS REFERENTE AO PERÍODO DE.: " & msk_data_periodo_inicial & " A " & msk_data_periodo_final, 0, 14, x_linha + 3.3, l_local
    ImprimeTexto "NOTAS REFERENTE AO PERÍODO DE.: " & msk_data_periodo_inicial & " A " & msk_data_periodo_final, 0, 14, x_linha + 3.315, l_local
    
    ImprimeTexto "APÓS O VENCIMENTO: PAGÁVEL APENAS NO " & l_nome_banco, 0, 14, x_linha + 3.7, l_local
    ImprimeTexto "APÓS O VENCIMENTO: PAGÁVEL APENAS NO " & l_nome_banco, 0, 14, x_linha + 3.715, l_local
    
    xJuros = Format(fValidaValor2(txt_valor_vencimento.Text) * 0.5 / 100, "00000000.00")
    ImprimeTexto "COBRAR R$ " & Format(xJuros, "##,###,##0.00") & " POR DIA DE ATRASO", 0, 10, x_linha + 4.1, l_local
    ImprimeTexto "COBRAR R$ " & Format(xJuros, "##,###,##0.00") & " POR DIA DE ATRASO", 0, 10, x_linha + 4.1093, l_local
    
    ImprimeTexto "PROTESTAR APÓS 5 DIAS DE VENCIMENTO", 0, 10, x_linha + 4.8, l_local
    ImprimeTexto "PROTESTAR APÓS 5 DIAS DE VENCIMENTO", 0, 10, x_linha + 4.813, l_local
    
    'ImprimeTexto "TX.BANC.", 11.7, 14, x_linha + 4.5, l_local
    
    
    ImprimeValor Format(dtcboCliente.BoundText, "##,###"), 0, 1, x_linha + 5.55, l_local
    ImprimeValor Format(dtcboCliente.BoundText, "##,###"), 0, 1, x_linha + 5.56, l_local
    ImprimeTexto dtcboCliente.Text, 1.1, 10, x_linha + 5.55, l_local
    ImprimeTexto dtcboCliente.Text, 1.1, 10, x_linha + 5.56, l_local
    
    If Cliente.LocalizarCodigo(Val(dtcboCliente.BoundText)) Then
        ImprimeTexto Trim(Cliente.Endereco) & " " & Cliente.Bairro, 1.1, 10.1, x_linha + 5.95, l_local
        ImprimeTexto Trim(Cliente.Endereco) & " " & Cliente.Bairro, 1.1, 10.1, x_linha + 5.965, l_local
        ImprimeTexto Mid(Cliente.CEP, 1, 2) & "." & Mid(Cliente.CEP, 3, 3) & "-" & Mid(Cliente.CEP, 6, 3), 10.2, 11.9, x_linha + 5.95, l_local
        ImprimeTexto Mid(Cliente.CEP, 1, 2) & "." & Mid(Cliente.CEP, 3, 3) & "-" & Mid(Cliente.CEP, 6, 3), 10.2, 11.9, x_linha + 5.965, l_local
        ImprimeTexto " - " & Cliente.Cidade & " " & Cliente.UF, 11.9, 17, x_linha + 5.95, l_local
        ImprimeTexto " - " & Cliente.Cidade & " " & Cliente.UF, 11.9, 17, x_linha + 5.965, l_local
        If Trim(Cliente.CGC) <> "" Then
            ImprimeTexto "CNPJ: " & fMascaraCNPJ(Cliente.CGC), 0, 5, x_linha + 6.35, l_local
            ImprimeTexto "I.E.: " & Cliente.InscricaoEstadual, 5, 10, x_linha + 6.35, l_local
        Else
            ImprimeTexto "CPF.: " & Cliente.CPF, 0, 5, x_linha + 6.35, l_local
            ImprimeTexto "R.G.: " & Cliente.Identidade, 5, 10, x_linha + 6.35, l_local
        End If
        ImprimeTexto "Fone: " & Cliente.Telefone, 10, 15, x_linha + 6.35, l_local
    End If
    Printer.EndDoc
End Sub
Private Sub SelecionaPaginaDuplicata()
Dim Impressora As Printer
Dim NumeroFormulario As Integer
Dim i As Integer
    On Error GoTo PrinterError
    i = 0
    NumeroFormulario = 47
    Printer.ScaleMode = 7
    Printer.PaperSize = NumeroFormulario
    Exit Sub
PrinterError:
    If i = 1 Then
        Exit Sub
    End If
    NumeroFormulario = 256
    i = i + 1
    Resume
End Sub
Private Sub BaixaNotaAbastecimento(ByVal pDataInicial As Date, ByVal pDataFinal As Date)
    Dim rsNotaAbastecimento As adodb.Recordset
    Dim xSQL As String
    
    On Error GoTo FileError
    
    Set rsNotaAbastecimento = New adodb.Recordset
    xSQL = ""
    xSQL = xSQL & "SELECT [Codigo do Cliente], [Data do Abastecimento], [Numero da Nota], Ordem, [Codigo do Produto2], [Periodo], Empresa"
    xSQL = xSQL & "  FROM Movimento_Nota_Abastecimento"
    'xSQL = xSQL & " WHERE Empresa = " & g_empresa
    xSQL = xSQL & " WHERE Empresa >= " & 1
    xSQL = xSQL & "   AND [Codigo do Cliente] = " & CLng(dtcboCliente.BoundText)
    xSQL = xSQL & "   AND [Data do Abastecimento] >= " & preparaData(pDataInicial)
    xSQL = xSQL & "   AND [Data do Abastecimento] <= " & preparaData(pDataFinal)
    xSQL = xSQL & " ORDER BY [Data do Abastecimento], [Periodo], [Numero da Nota], Ordem, [Codigo do Produto2]"
    Set rsNotaAbastecimento = Conectar.RsConexao(xSQL)
    With rsNotaAbastecimento
        If .RecordCount > 0 Then
            Do Until .EOF
                If MovNotaAbastecimento.LocalizarCodigo(rsNotaAbastecimento("Empresa").Value, rsNotaAbastecimento("Codigo do Cliente").Value, rsNotaAbastecimento("Data do Abastecimento").Value, rsNotaAbastecimento("Numero da Nota").Value, rsNotaAbastecimento("Ordem").Value, rsNotaAbastecimento("Codigo do Produto2").Value, rsNotaAbastecimento("Periodo").Value) Then
                    MovBaixaNotaAbastecimento.CodigoCliente = MovNotaAbastecimento.CodigoCliente
                    MovBaixaNotaAbastecimento.DataPagamento = msk_data_emissao.Text
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
                    MovBaixaNotaAbastecimento.NumeroMovimentoCaixaBaixa = 0
                    MovBaixaNotaAbastecimento.BaixadoPelaDuplicata = True
                    MovBaixaNotaAbastecimento.NumeroIlha = MovNotaAbastecimento.NumeroIlha
                    MovBaixaNotaAbastecimento.Origem = MovNotaAbastecimento.Origem
                    MovBaixaNotaAbastecimento.DataBaixa = CDate(msk_data_emissao.Text)
                    MovBaixaNotaAbastecimento.NumeroCupom = MovNotaAbastecimento.NumeroCupom
                    MovBaixaNotaAbastecimento.DataConferencia = MovNotaAbastecimento.DataConferencia
                    MovBaixaNotaAbastecimento.NumeroDuplicata = CLng(txt_numero_documento.Text)
                    MovBaixaNotaAbastecimento.KM = MovNotaAbastecimento.KM
                    If MovBaixaNotaAbastecimento.Incluir Then
                        If Not MovNotaAbastecimento.Excluir(rsNotaAbastecimento("Empresa").Value, rsNotaAbastecimento("Codigo do Cliente").Value, rsNotaAbastecimento("Data do Abastecimento").Value, rsNotaAbastecimento("Numero da Nota").Value, rsNotaAbastecimento("Ordem").Value, rsNotaAbastecimento("Codigo do Produto2").Value, rsNotaAbastecimento("Periodo").Value) Then
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
    Exit Sub
    
FileError:
    'ErroArquivo tbl_baixa_nota_abastecimento.Name, "Baixaa"
    Exit Sub
End Sub
Private Sub EstornaBaixaNotaAbastecimento(ByVal pNumeroDocumento As Long, ByVal pDataInicial As Date, ByVal pDataFinal As Date)
    
    Dim rsBaixaNotaAbastecimento As adodb.Recordset
    Dim xSQL As String
    
    On Error GoTo FileError
    
    xSQL = "UPDATE Movimento_Nota_Abastecimento"
    xSQL = xSQL & "   SET [Baixado Pela Duplicata] = " & preparaBooleano(False)
    xSQL = xSQL & " WHERE [Codigo do Cliente] = " & CLng(dtcboCliente.BoundText)
    xSQL = xSQL & "   AND [Data do Abastecimento] >= " & preparaData(pDataInicial)
    xSQL = xSQL & "   AND [Data do Abastecimento] <= " & preparaData(pDataFinal)
    Conectar.ExecutaSql (xSQL)
    
    Set rsBaixaNotaAbastecimento = New adodb.Recordset
    xSQL = ""
    xSQL = xSQL & "SELECT [Codigo do Cliente], [Data do Abastecimento], [Numero da Nota], Ordem, [Codigo do Produto2], [Periodo], Empresa"
    xSQL = xSQL & "  FROM Baixa_Nota_Abastecimento"
    'xSQL = xSQL & " WHERE Empresa = " & g_empresa
    xSQL = xSQL & " WHERE Empresa >= " & 1
    xSQL = xSQL & "   AND [Codigo do Cliente] = " & CLng(dtcboCliente.BoundText)
    xSQL = xSQL & "   AND [Data do Abastecimento] >= " & preparaData(pDataInicial)
    xSQL = xSQL & "   AND [Data do Abastecimento] <= " & preparaData(pDataFinal)
    xSQL = xSQL & "   AND [Numero da Duplicata] = " & pNumeroDocumento
    xSQL = xSQL & " ORDER BY [Data do Abastecimento], [Periodo], [Numero da Nota], Ordem, [Codigo do Produto2]"
    Set rsBaixaNotaAbastecimento = Conectar.RsConexao(xSQL)
    If rsBaixaNotaAbastecimento.RecordCount > 0 Then
        Do Until rsBaixaNotaAbastecimento.EOF
            If MovBaixaNotaAbastecimento.LocalizarCodigo(rsBaixaNotaAbastecimento("Empresa").Value, rsBaixaNotaAbastecimento("Codigo do Cliente").Value, rsBaixaNotaAbastecimento("Data do Abastecimento").Value, rsBaixaNotaAbastecimento("Numero da Nota").Value, rsBaixaNotaAbastecimento("Ordem").Value, rsBaixaNotaAbastecimento("Codigo do Produto2").Value, rsBaixaNotaAbastecimento("Periodo").Value) Then
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
                MovNotaAbastecimento.NumeroCupom = MovBaixaNotaAbastecimento.NumeroCupom
                MovNotaAbastecimento.ValorDescontoUnitario = MovBaixaNotaAbastecimento.ValorDescontoUnitario
                MovNotaAbastecimento.NumeroMovimentoCaixa = MovBaixaNotaAbastecimento.NumeroMovimentoCaixa
                MovNotaAbastecimento.BaixadoPelaDuplicata = MovBaixaNotaAbastecimento.BaixadoPelaDuplicata
                MovNotaAbastecimento.NumeroIlha = MovBaixaNotaAbastecimento.NumeroIlha
                MovNotaAbastecimento.Origem = MovBaixaNotaAbastecimento.Origem
                MovNotaAbastecimento.DataConferencia = MovBaixaNotaAbastecimento.DataConferencia
                MovNotaAbastecimento.KM = MovBaixaNotaAbastecimento.KM
                If MovNotaAbastecimento.Incluir Then
                    If Not MovBaixaNotaAbastecimento.Excluir(rsBaixaNotaAbastecimento("Empresa").Value, rsBaixaNotaAbastecimento("Codigo do Cliente").Value, rsBaixaNotaAbastecimento("Data do Abastecimento").Value, rsBaixaNotaAbastecimento("Numero da Nota").Value, rsBaixaNotaAbastecimento("Ordem").Value, rsBaixaNotaAbastecimento("Codigo do Produto2").Value, rsBaixaNotaAbastecimento("Periodo").Value) Then
                        MsgBox "Não foi possível excluir este registro!", vbInformation, "Erro de Integridade."
                    End If
                Else
                    MsgBox "Não foi possível incluir este registro!", vbInformation, "Erro de Integridade."
                End If
            Else
                MsgBox "Não foi possível localizar este registro!", vbInformation, "Erro de Integridade."
            End If
            rsBaixaNotaAbastecimento.MoveNext
        Loop
    End If
    rsBaixaNotaAbastecimento.Close
    Set rsBaixaNotaAbastecimento = Nothing
    Exit Sub

FileError:
    MsgBox "Erro ao extornar notas de abastecimento.", vbInformation, "Rotina: EstornaBaixaNotaAbastecimento"
    Exit Sub
End Sub
Public Function GravaSgpNetCadastroIni(ByVal pNomePrograma As String) As Boolean
    Dim xString As String
    Dim xArquivoTmp As String
    Dim retval As Long
    Dim xNomeArquivo As String
        
    On Error GoTo FileError
    
    xNomeArquivo = "C:\Cerrado.Net\SgpNet\SgpNetTemporario.ini"
    GravaSgpNetCadastroIni = False
    'lNomeArquivo = gDrive & Mid(gDiretorioData, 1, Len(gDiretorioData) - 5) & "sgp_cadastro.ini"
    retval = Len(xNomeArquivo)
    xArquivoTmp = Mid(xNomeArquivo, 1, retval - 3) & "tmp"
    If gArqTxt.FileExists(xArquivoTmp) Then
        Call gArqTxt.DeleteFile(xArquivoTmp, True)
    End If
    If gArqTxt.FileExists(xNomeArquivo) Then
        Call gArqTxt.DeleteFile(xNomeArquivo, True)
    End If
    
    Set gArquivoTMP = gArqTxt.CreateTextFile(xArquivoTmp)
    gArquivoTMP.WriteLine ("[Empresa]")
    gArquivoTMP.WriteLine ("Empresa=" & Format(g_empresa, "000"))
    gArquivoTMP.WriteLine ("NomeEmpresa=" & g_nome_empresa)
    gArquivoTMP.WriteLine ("CidadeEmpresa=" & g_cidade_empresa)
    
    gArquivoTMP.WriteLine (" ")
    gArquivoTMP.WriteLine ("[Programa]")
    gArquivoTMP.WriteLine ("Programa=" & pNomePrograma)
    
    gArquivoTMP.WriteLine (" ")
    gArquivoTMP.WriteLine ("[Usuario]")
    gArquivoTMP.WriteLine ("Usuario=" & CStr(g_usuario))
    gArquivoTMP.WriteLine ("NomeUsuario=" & g_nome_usuario)
    gArquivoTMP.WriteLine ("NivelAcessoUsuario=" & CStr(g_nivel_acesso))
    
    gArquivoTMP.WriteLine (" ")
    gArquivoTMP.WriteLine ("[Outras]")
    gArquivoTMP.WriteLine ("DataDef=" & CStr(g_data_def))
    gArquivoTMP.WriteLine ("FlagLmc=" & CStr(g_lmc))
    gArquivoTMP.WriteLine ("ImpressoraMatricial=" & CStr(g_impressora_matricial))
    gArquivoTMP.WriteLine ("CaixaUnificado=" & CStr(g_caixa_unificado))
    
    gArquivoTMP.WriteLine (" ")
    gArquivoTMP.WriteLine ("[ContaBancaria]")
    gArquivoTMP.WriteLine ("ContaBancaria=" & CStr(g_conta_bancaria))
    gArquivoTMP.WriteLine ("NomeContaBancaria=" & CStr(g_nome_conta))
    
    gArquivoTMP.WriteLine (" ")
    gArquivoTMP.WriteLine ("[Liberacao]")
    gArquivoTMP.WriteLine ("EmpresaInicial=" & CStr(g_cfg_empresa_i))
    gArquivoTMP.WriteLine ("EmpresaFinal=" & CStr(g_cfg_empresa_f))
    gArquivoTMP.WriteLine ("DataInicial=" & CStr(g_cfg_data_i))
    gArquivoTMP.WriteLine ("DataFinal=" & CStr(g_cfg_data_f))
    gArquivoTMP.WriteLine ("PeriodoInicial=" & CStr(g_cfg_periodo_i))
    gArquivoTMP.WriteLine ("PeriodoFinal=" & CStr(g_cfg_periodo_f))
    
    gArquivoTMP.WriteLine (" ")
    gArquivoTMP.WriteLine ("[Automacao]")
    gArquivoTMP.WriteLine ("Automacao=" & CStr(g_automacao))
    
    gArquivoTMP.WriteLine (" ")
    gArquivoTMP.WriteLine ("[String]")
    gArquivoTMP.WriteLine ("String de Chamada=" & gStringChamada)
    
    gArquivoTMP.Close
    
    
    
    Call gArqTxt.MoveFile(xArquivoTmp, xNomeArquivo)

    GravaSgpNetCadastroIni = True
    Exit Function

FileError:
    MsgBox "Erro ao gravar Sgp_Cadastro.ini!" & Chr(10) & Error, vbInformation, "Erro Interno!"
End Function

