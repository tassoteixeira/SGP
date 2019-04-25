VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form lst_falta_caixa 
   Caption         =   "Emissão da Falta de Caixa / Vales de Funcionários"
   ClientHeight    =   3555
   ClientLeft      =   3990
   ClientTop       =   2010
   ClientWidth     =   6795
   Icon            =   "lst_falta_caixa.frx":0000
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   Picture         =   "lst_falta_caixa.frx":030A
   ScaleHeight     =   3555
   ScaleWidth      =   6795
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1140
      Picture         =   "lst_falta_caixa.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Visualiza falta de caixa/vale."
      Top             =   2580
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3000
      Picture         =   "lst_falta_caixa.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Imprime falta de caixa/vale."
      Top             =   2580
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4860
      Picture         =   "lst_falta_caixa.frx":3074
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   2580
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   2355
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6555
      Begin VB.ComboBox cboFuncionario 
         Height          =   315
         Left            =   2100
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1920
         Width           =   4335
      End
      Begin VB.ComboBox cboTipoMovimento 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1500
         Width           =   1995
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   5940
         Picture         =   "lst_falta_caixa.frx":4706
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2760
         Picture         =   "lst_falta_caixa.frx":59E0
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2760
         Picture         =   "lst_falta_caixa.frx":6CBA
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txt_funcionario 
         Height          =   315
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   17
         Top             =   1920
         Width           =   375
      End
      Begin VB.ComboBox cbo_periodo_f 
         Height          =   315
         Left            =   5940
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1080
         Width           =   495
      End
      Begin VB.ComboBox cbo_periodo_i 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1080
         Width           =   495
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
      Begin VB.Label Label7 
         Caption         =   "&Tipo do Movimento"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1500
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "&Período Final"
         Height          =   255
         Left            =   3840
         TabIndex        =   12
         Top             =   1080
         Width           =   1515
      End
      Begin VB.Label Label4 
         Caption         =   "&Funcionário"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1920
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "Período &Inicial"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "&Data Inicial"
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   660
         Width           =   1515
      End
      Begin VB.Label Label2 
         Caption         =   "D&ata Final"
         Height          =   315
         Left            =   3840
         TabIndex        =   7
         Top             =   660
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Data de Emissão"
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1515
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   2820
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "lst_falta_caixa"
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
Dim lSQL As String
Private Funcionario As New cFuncionario
Private rsMovFaltaCaixa As New adodb.Recordset

Private Const NIVEL_ACESSO_DIGITACAO As Integer = 5

Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set rsMovFaltaCaixa = Nothing
End Sub
Private Sub PreencheCboFuncionario()
     Dim xSQL As String
    Dim rsTabela As New adodb.Recordset

    cboFuncionario.Clear
    cboFuncionario.AddItem "Todos os Funcionários"
    cboFuncionario.ItemData(cboFuncionario.NewIndex) = 0
    'Prepara SQL
    xSQL = "SELECT Codigo, Nome"
    xSQL = xSQL & "  FROM Funcionario"
    xSQL = xSQL & " WHERE Empresa = " & g_empresa
    xSQL = xSQL & "   AND Situacao = " & preparaTexto("A")
    xSQL = xSQL & "   AND [Periodo] < 5"
    xSQL = xSQL & " ORDER BY Nome"
    'Abre RecordSet
    Set rsTabela = New adodb.Recordset
    Set rsTabela = Conectar.RsConexao(xSQL)
    'Verifica movimento
    If rsTabela.RecordCount > 0 Then
        Do Until rsTabela.EOF
            cboFuncionario.AddItem rsTabela("Nome").Value
            cboFuncionario.ItemData(cboFuncionario.NewIndex) = rsTabela("Codigo").Value
            rsTabela.MoveNext
        Loop
    End If
    
    
    If rsTabela.State = 1 Then
        rsTabela.Close
    End If
End Sub
    Private Function PreparaTelaUsuarioLogado(ByVal pCodigoFuncionario As Integer) As Boolean
        PreparaTelaUsuarioLogado = True
        
        Dim xFuncionario As New CadastroDLL.cFuncionario
        cboFuncionario.ListIndex = -1
        txt_funcionario.Text = ""
        cboFuncionario.Enabled = True
        txt_funcionario.Enabled = True
        cmd_imprimir.Enabled = True
        cmd_visualizar.Enabled = True

        If g_nivel_acesso >= NIVEL_ACESSO_DIGITACAO Then

            If pCodigoFuncionario <= 0 Then
                If xFuncionario.LocalizarFuncionarioDoUsuario(g_usuario, g_empresa) Then
                    If SelecionaFuncionarioNaCombo(xFuncionario.Codigo) Then
                        txt_funcionario.Text = CStr(xFuncionario.Codigo)
                        cboFuncionario.Enabled = False
                        txt_funcionario.Enabled = False
                    Else
                        MsgBox "O usuário logado não tem permissão para esta funcionalidade.", vbInformation, "Atenção!"
                        cmd_imprimir.Enabled = False
                        cmd_visualizar.Enabled = False
                        PreparaTelaUsuarioLogado = False
                    End If
                Else
                    MsgBox "O usuário logado não tem permissão para esta funcionalidade.", vbInformation, "Atenção!"
                    cmd_imprimir.Enabled = False
                    cmd_visualizar.Enabled = False
                    PreparaTelaUsuarioLogado = False
                End If
            Else
                If xFuncionario.LocalizarCodigo(g_empresa, pCodigoFuncionario) Then
                    If SelecionaFuncionarioNaCombo(xFuncionario.Codigo) Then
                        txt_funcionario.Text = CStr(xFuncionario.Codigo)
                        cboFuncionario.Enabled = False
                        txt_funcionario.Enabled = False
                    Else
                        MsgBox "O usuário logado não tem permissão para esta funcionalidade.", vbInformation, "Atenção!"
                        cmd_imprimir.Enabled = False
                        cmd_visualizar.Enabled = False
                        PreparaTelaUsuarioLogado = False
                    End If
                Else
                    MsgBox "O usuário logado não tem permissão para esta funcionalidade.", vbInformation, "Atenção!"
                    cmd_imprimir.Enabled = False
                    cmd_visualizar.Enabled = False
                    PreparaTelaUsuarioLogado = False
                End If
            End If
        End If
    End Function

Private Sub PreencheCboPeriodo()
    cbo_periodo_i.Clear
    cbo_periodo_f.Clear
    cbo_periodo_i.AddItem 1
    cbo_periodo_f.AddItem 1
    cbo_periodo_f.ItemData(cbo_periodo_f.NewIndex) = 1
    cbo_periodo_i.ItemData(cbo_periodo_i.NewIndex) = 1
    cbo_periodo_i.AddItem 2
    cbo_periodo_f.AddItem 2
    cbo_periodo_i.ItemData(cbo_periodo_i.NewIndex) = 2
    cbo_periodo_f.ItemData(cbo_periodo_f.NewIndex) = 2
    cbo_periodo_i.AddItem 3
    cbo_periodo_f.AddItem 3
    cbo_periodo_i.ItemData(cbo_periodo_i.NewIndex) = 3
    cbo_periodo_f.ItemData(cbo_periodo_f.NewIndex) = 3
    cbo_periodo_i.AddItem 4
    cbo_periodo_f.AddItem 4
    cbo_periodo_i.ItemData(cbo_periodo_i.NewIndex) = 4
    cbo_periodo_f.ItemData(cbo_periodo_f.NewIndex) = 4
End Sub
Private Sub PreencheCboTipoMovimento()
    cboTipoMovimento.Clear
    cboTipoMovimento.AddItem "Geral"
    cboTipoMovimento.ItemData(cboTipoMovimento.NewIndex) = 0
    cboTipoMovimento.AddItem "Falta de Caixa"
    cboTipoMovimento.ItemData(cboTipoMovimento.NewIndex) = 1
    cboTipoMovimento.AddItem "Sobra de Caixa"
    cboTipoMovimento.ItemData(cboTipoMovimento.NewIndex) = 2
    cboTipoMovimento.AddItem "Vale"
    cboTipoMovimento.ItemData(cboTipoMovimento.NewIndex) = 3
End Sub
Private Sub ZeraVariaveis()
    lLinha = 0
    lPagina = 0
    lTotal = 0
End Sub
Private Sub PreparaDatas()
    If Month(msk_data.Text) = 1 Then
        msk_data_i.Text = CDate("26/12/" & Year(msk_data.Text) - 1)
        msk_data_f.Text = CDate("25/" & Month(msk_data.Text) & "/" & Year(msk_data.Text))
    Else
        msk_data_i.Text = CDate("26/" & Month(msk_data.Text) - 1 & "/" & Year(msk_data.Text))
        msk_data_f.Text = CDate("25/" & Month(msk_data.Text) & "/" & Year(msk_data.Text))
    End If
End Sub
Private Sub Relatorio()
    ZeraVariaveis
    
    'Prepara SQL
    If cboFuncionario.ItemData(cboFuncionario.ListIndex) > 0 Then
        lSQL = "SELECT Data, Periodo, Valor, Observacao, [Tipo de Movimento]"
        lSQL = lSQL & " FROM Movimento_Falta_Caixa"
        lSQL = lSQL & " WHERE Empresa = " & g_empresa
        lSQL = lSQL & " AND Data >= " & preparaData(CDate(msk_data_i.Text))
        lSQL = lSQL & " AND Data <= " & preparaData(CDate(msk_data_f.Text))
        lSQL = lSQL & " AND Periodo >= " & preparaTexto(cbo_periodo_i.Text)
        lSQL = lSQL & " AND Periodo <= " & preparaTexto(cbo_periodo_f.Text)
        If cboTipoMovimento.ListIndex = 1 Then
            lSQL = lSQL & " AND [Tipo de Movimento] = " & preparaTexto("F")
        ElseIf cboTipoMovimento.ListIndex = 2 Then
            lSQL = lSQL & " AND [Tipo de Movimento] = " & preparaTexto("S")
        ElseIf cboTipoMovimento.ListIndex = 3 Then
            lSQL = lSQL & " AND [Tipo de Movimento] = " & preparaTexto("V")
        End If
        lSQL = lSQL & " AND [Codigo do Funcionario] = " & Val(txt_funcionario.Text)
        lSQL = lSQL & " ORDER BY Data, Periodo"
    Else
        lSQL = "SELECT Funcionario.Nome, Movimento_Falta_Caixa.[Codigo do Funcionario], SUM(Movimento_Falta_Caixa.Valor) AS Total"
        lSQL = lSQL & " FROM Movimento_Falta_Caixa, Funcionario"
        lSQL = lSQL & " WHERE Movimento_Falta_Caixa.Empresa = " & g_empresa
        lSQL = lSQL & " AND Movimento_Falta_Caixa.Data >= " & preparaData(CDate(msk_data_i.Text))
        lSQL = lSQL & " AND Movimento_Falta_Caixa.Data <= " & preparaData(CDate(msk_data_f.Text))
        lSQL = lSQL & " AND Movimento_Falta_Caixa.Periodo >= " & preparaTexto(cbo_periodo_i.Text)
        lSQL = lSQL & " AND Movimento_Falta_Caixa.Periodo <= " & preparaTexto(cbo_periodo_f.Text)
        If cboTipoMovimento.ListIndex = 1 Then
            lSQL = lSQL & " AND Movimento_Falta_Caixa.[Tipo de Movimento] = " & preparaTexto("F")
        ElseIf cboTipoMovimento.ListIndex = 2 Then
            lSQL = lSQL & " AND Movimento_Falta_Caixa.[Tipo de Movimento] = " & preparaTexto("S")
        ElseIf cboTipoMovimento.ListIndex = 3 Then
            lSQL = lSQL & " AND Movimento_Falta_Caixa.[Tipo de Movimento] = " & preparaTexto("V")
        End If
        lSQL = lSQL & " AND Movimento_Falta_Caixa.[Codigo do Funcionario] = Funcionario.Codigo"
        lSQL = lSQL & " AND Funcionario.Empresa = " & g_empresa
        lSQL = lSQL & " GROUP BY Funcionario.Nome, Movimento_Falta_Caixa.[Codigo do Funcionario]"
        lSQL = lSQL & " ORDER BY Funcionario.Nome"
    End If
    
    'Abre RecordSet
    Set rsMovFaltaCaixa = New adodb.Recordset
    Set rsMovFaltaCaixa = Conectar.RsConexao(lSQL)
    
    
    'Verifica movimento
    If rsMovFaltaCaixa.RecordCount > 0 Then
        ImpDados
    End If
    If rsMovFaltaCaixa.State = 1 Then
        rsMovFaltaCaixa.Close
    End If
'    adodc_funcionario.Recordset.Find ("Codigo = " & Val(dtcbo_funcionario.BoundText))
'    If Not adodc_funcionario.Recordset.EOF Then
'        adodc_funcionario.Recordset.MoveNext
'        If Not adodc_funcionario.Recordset.EOF Then
'            dtcbo_funcionario.BoundText = adodc_funcionario.Recordset!Codigo
'        End If
'    End If
    If cboFuncionario.Enabled = True Then
       cboFuncionario.SetFocus
    End If
End Sub
Private Function SelecionaFuncionarioNaCombo(ByVal pCodigoFuncionario As Integer) As Boolean
    Dim i As Integer
    SelecionaFuncionarioNaCombo = False
    
    If pCodigoFuncionario > 0 Then
        If Funcionario.LocalizarCodigo(g_empresa, pCodigoFuncionario) Then
            If Funcionario.Situacao = "I" Then
                MsgBox "O funcionário " & Trim(Funcionario.Nome) & " está inativo.", vbInformation, "Atenção!"
                Exit Function
            Else
                cboFuncionario.ListIndex = -1
                For i = 0 To cboFuncionario.ListCount - 1
                    If cboFuncionario.ItemData(i) = pCodigoFuncionario Then
                        cboFuncionario.ListIndex = i
                        SelecionaFuncionarioNaCombo = True
                        Exit For
                    End If
                Next
            End If
        Else
            MsgBox "Funcionário não cadastrado.", vbInformation, "Atenção!"
            Exit Function
        End If
    End If
End Function

Private Sub ImpDados()
    LoopMovimentoFaltaCaixa
    If lPagina > 0 Then
        ImpTotal
        BioImprime "@@Printer.EndDoc"
        BioFechaImprime
        g_string = lLocal & lNomeArquivo & "|@|Relatório de Falta de Caixa/Vales de Funcionários|@|"
        frm_preview.Show 1
    End If
End Sub
Private Sub LoopMovimentoFaltaCaixa()
    'loop movimento de falta de caixa
    Dim x_linha As String
    Do Until rsMovFaltaCaixa.EOF
        If lPagina = 0 Then
            ImpCab
        End If
        If lLinha >= 60 Then
            x_linha = "+----------+---+--------+-------------+----------------------------------------+"
            Mid(x_linha, 39, 22) = " Cerrado Informática. "
            BioImprime "@Printer.Print " & x_linha
            BioImprime "@@Printer.NewPage"
            ImpCab
        End If
        If cboFuncionario.ItemData(cboFuncionario.ListIndex) > 0 Then
            Call ImpDet
            lTotal = lTotal + rsMovFaltaCaixa("Valor").Value
        Else
            Call ImpDet2
            lTotal = lTotal + rsMovFaltaCaixa("Total").Value
        End If
        rsMovFaltaCaixa.MoveNext
    Loop
End Sub
Private Sub ImpDet()
    Dim x_linha As String
    Dim i As Integer
    x_linha = "|          |   |        |             |                                        |"
    Mid(x_linha, 2, 10) = Format(rsMovFaltaCaixa("Data").Value, "dd/mm/yyyy")
    Mid(x_linha, 14, 1) = rsMovFaltaCaixa("Periodo").Value
    If rsMovFaltaCaixa("Tipo de Movimento").Value = "F" Then
        Mid(x_linha, 17, 8) = "FALTA CX"
    ElseIf rsMovFaltaCaixa("Tipo de Movimento").Value = "S" Then
        Mid(x_linha, 17, 8) = "SOBRA CX"
    ElseIf rsMovFaltaCaixa("Tipo de Movimento").Value = "V" Then
        Mid(x_linha, 17, 8) = "VALE    "
    End If
    i = Len(Format(rsMovFaltaCaixa("Valor").Value, "####,##0.00"))
    Mid(x_linha, 27 + 11 - i, i) = Format(rsMovFaltaCaixa("Valor").Value, "####,##0.00")
    Mid(x_linha, 40, 40) = rsMovFaltaCaixa("Observacao").Value
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & x_linha
    lLinha = lLinha + 1
End Sub
Private Sub ImpDet2()
    Dim x_linha As String
    Dim i As Integer
    
    x_linha = "|        |                                               |                     |"
    
'Nome, Codigo do Funcionario, Total
    
    i = Len(Format(rsMovFaltaCaixa("Codigo do Funcionario").Value, "##,##0"))
    Mid(x_linha, 3 + 6 - i, i) = Format(rsMovFaltaCaixa("Codigo do Funcionario").Value, "##,##0")
    Mid(x_linha, 12, 40) = rsMovFaltaCaixa("Nome").Value
    i = Len(Format(rsMovFaltaCaixa("Total").Value, "####,##0.00"))
    Mid(x_linha, 68 + 11 - i, i) = Format(rsMovFaltaCaixa("Total").Value, "####,##0.00")
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & x_linha
    lLinha = lLinha + 1
End Sub
Private Sub ImpTotal()
    Dim y_local As Single
    Dim x_linha As String
    Dim i As Integer
    
    If cboFuncionario.ItemData(cboFuncionario.ListIndex) > 0 Then
        x_linha = "+----------+---+--------+-------------+----------------------------------------+"
        BioImprime "@Printer.Print " & x_linha
        x_linha = "|           *** TOTAL.: |             |                                        |"
        i = Len(Format(lTotal, "####,##0.00"))
        Mid(x_linha, 27 + 11 - i, i) = Format(lTotal, "####,##0.00")
        If lTotal > 0 Then
            Mid(x_linha, 40, 40) = "DESCONTAR DO FUNCIONÁRIO"
        ElseIf lTotal < 0 Then
            Mid(x_linha, 40, 40) = "RESTITUIR AO FUNCIONÁRIO"
        End If
        BioImprime "@@Printer.FontName = Draft 10cpi"
        BioImprime "@@y_local = Printer.CurrentY"
        BioImprime "@@ImprimeTexto " & Chr(34) & "  " & Chr(34) & ", 1, 2, 2, 1"
        BioImprime "@@Printer.CurrentY = y_local"
        BioImprime "@@Printer.FontBold = True"
        BioImprime "@Printer.Print " & x_linha
        BioImprime "@@Printer.CurrentY = y_local"
        BioImprime "@@Printer.Print " & "  "
        BioImprime "@@Printer.FontBold = False"
        x_linha = "+-----------------------+-------------+----------------------------------------+"
        BioImprime "@Printer.Print " & x_linha
        BioImprime "@Printer.Print " & "|                                                                              |"
        BioImprime "@Printer.Print " & "|                                                                              |"
        BioImprime "@Printer.Print " & "|     Foi feita a conferência do caixa referente a movimentação  de  combustí- |"
        BioImprime "@Printer.Print " & "| veis, conforme planilha de movimentação das bombas  deste  período,  o  qual |"
        BioImprime "@Printer.Print " & "| reconheço o total da falta de caixa, estando ciente que será  descontado  em |"
        x_linha = "| meu salário do mês de          /      conforme  cláusula  10a  da  Convenção |"
        Mid(x_linha, 25, 9) = Format(msk_data_f, "mmmm")
        Mid(x_linha, 35, 4) = Format(msk_data_f, "yyyy")
        BioImprime "@Printer.Print " & x_linha
        BioImprime "@Printer.Print " & "| Coletiva do Trabalho em conjunto com o Regulamento Interno do Posto.         |                                                                   |"
        BioImprime "@Printer.Print " & "|                                                                              |"
        BioImprime "@Printer.Print " & "|     Declaro ainda que o erro que originou a falta  de  caixa  foi  de  minha |"
        BioImprime "@Printer.Print " & "| responsabilidade.                                                            |"
        BioImprime "@Printer.Print " & "|                                                                              |"
        BioImprime "@Printer.Print " & "|                                                                              |"
        BioImprime "@Printer.Print " & "|                               Goiânia-GO, ____ de ______________ de _______. |"
        BioImprime "@Printer.Print " & "|                                                                              |"
        BioImprime "@Printer.Print " & "|                                                                              |"
        BioImprime "@Printer.Print " & "|                  __________________________________________                  |"
        x_linha = "|                                                                              |"
        i = Val((40 - Len(Trim(cboFuncionario.Text))) / 2)
        Mid(x_linha, 21 + i, Len(Trim(cboFuncionario.Text))) = Trim(cboFuncionario.Text)
        BioImprime "@Printer.Print " & x_linha
        BioImprime "@Printer.Print " & "|                                                                              |"
        x_linha = "+------------------------------------------------------------------------------+"
        Mid(x_linha, 39, 22) = " Cerrado Informática. "
        BioImprime "@Printer.Print " & x_linha
        BioImprime "@Printer.Print " & "  "
        BioImprime "@@Printer.FontBold = True"
        BioImprime "@Printer.Print " & "OBS.: O FUNCIONÁRIO TEM 2 DIAS PARA RECLAMAR ALGUMA DESSAS FALTAS!"
        BioImprime "@@Printer.FontBold = False"
    Else
        x_linha = "+--------+-----------------------------------------------+---------------------+"
        BioImprime "@Printer.Print " & x_linha
        BioImprime "@@Printer.FontBold = True"
        x_linha = "|                                              *** TOTAL |                     |"
        i = Len(Format(lTotal, "####,##0.00"))
        Mid(x_linha, 68 + 11 - i, i) = Format(lTotal, "####,##0.00")
        BioImprime "@Printer.Print " & x_linha
        BioImprime "@@Printer.FontBold = False"
        x_linha = "+--------------------------------------------------------+---------------------+"
        Mid(x_linha, 5, 22) = " Cerrado Informática. "
        BioImprime "@Printer.Print " & x_linha
    End If
    BioImprime "@Printer.Print " & "  "
End Sub
Private Sub ImpCab()
    Dim x_linha As String
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
    BioImprime "@Printer.Print " & "+------------------------------------------------------------------------------+"
    BioImprime "@@Printer.FontBold = True"
    x_linha = "|                                                                  Página,     |"
    Mid(x_linha, 3, 40) = g_nome_empresa
    Mid(x_linha, 76, 3) = Format(lPagina, "000")
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontBold = False"
    x_linha = "| RELAÇÃO DE FALTA DE CAIXA                                 CIDADE,            |"
    i = Len(g_cidade_empresa)
    Mid(x_linha, 37 + 30 - i, i) = g_cidade_empresa
    Mid(x_linha, 69, 10) = msk_data.Text
    BioImprime "@Printer.Print " & x_linha
    x_linha = "| REFERENTE AO PERIODO DE.: __/__/____ A __/__/____                            |"
    Mid(x_linha, 29, 10) = msk_data_i.Text
    Mid(x_linha, 42, 10) = msk_data_f.Text
    BioImprime "@Printer.Print " & x_linha
    x_linha = "| CAIXA INICIAL...........: X    CAIXA FINAL..: X                              |"
    Mid(x_linha, 29, 1) = cbo_periodo_i.Text
    Mid(x_linha, 49, 1) = cbo_periodo_f.Text
    BioImprime "@Printer.Print " & x_linha
    x_linha = "| FUNCIONARIO.............:                                                    |"
    If cboFuncionario.ItemData(cboFuncionario.ListIndex) > 0 Then
        Mid(x_linha, 29, 3) = Format(Funcionario.Codigo, "000")
        Mid(x_linha, 33, Len(Funcionario.Nome)) = Funcionario.Nome
    Else
        Mid(x_linha, 29, 3) = Format(Val(txt_funcionario.Text), "000")
        Mid(x_linha, 33, 40) = cboFuncionario.Text
    End If
    BioImprime "@Printer.Print " & x_linha
    x_linha = "| TIPO DE MOVIMENTO.......:                                                    |"
    Mid(x_linha, 29, 20) = cboTipoMovimento.Text
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontName = Draft 10cpi"
    If cboFuncionario.ItemData(cboFuncionario.ListIndex) > 0 Then
        BioImprime "@Printer.Print " & "+----------+---+--------+-------------+----------------------------------------+"
        BioImprime "@Printer.Print " & "|   DATA   |PER|TIPO MOV|  V A L O R  | OBSERVACAO                             |"
        BioImprime "@Printer.Print " & "+----------+---+--------+-------------+----------------------------------------+"
    Else
        BioImprime "@Printer.Print " & "+--------+-----------------------------------------------+---------------------+"
        BioImprime "@Printer.Print " & "| CODIGO | NOME DO FUNCIONARIO                           |   VALOR     TOTAL   |"
        BioImprime "@Printer.Print " & "+--------+-----------------------------------------------+---------------------+"
    End If
End Sub
Private Sub cbo_periodo_f_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cboTipoMovimento.SetFocus
    End If
End Sub
Private Sub cbo_periodo_i_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_periodo_f.SetFocus
    End If
End Sub
Private Sub cboFuncionario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data.SetFocus
    End If
End Sub
Private Sub cboFuncionario_LostFocus()
    If cboFuncionario.Text <> "" Then
        txt_funcionario.Text = cboFuncionario.ItemData(cboFuncionario.ListIndex)
        txt_funcionario_LostFocus
        If cmd_imprimir.Enabled = True Then
            cmd_imprimir.SetFocus
        Else
            cmd_sair.SetFocus
        End If
    End If
End Sub
Private Sub cboTipoMovimento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt_funcionario.Enabled = True Then
            txt_funcionario.SetFocus
        ElseIf cmd_imprimir.Enabled = True Then
            cmd_imprimir.SetFocus
        Else
            cmd_sair.SetFocus
        End If
    End If
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i.Text = RetiraGString(1)
        msk_data_f.Text = RetiraGString(2)
        cbo_periodo_i.SetFocus
    Else
        msk_data.Text = RetiraGString(1)
        msk_data_i.SetFocus
    End If
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
    cbo_periodo_i.SetFocus
End Sub
Private Sub cmd_data_i_Click()
    g_string = msk_data_i.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i.Text = RetiraGString(1)
        msk_data_f.Text = RetiraGString(2)
        cbo_periodo_i.SetFocus
    Else
        msk_data_i.Text = RetiraGString(1)
        msk_data_f.SetFocus
    End If
    g_string = ""
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
    ElseIf Not IsDate(msk_data_i.Text) Then
        MsgBox "Informe a data inicial.", vbInformation, "Atenção!"
        msk_data_i.SetFocus
    ElseIf Not IsDate(msk_data_f.Text) Then
        MsgBox "Informe a data final.", vbInformation, "Atenção!"
        msk_data_f.SetFocus
    ElseIf CDate(msk_data_f.Text) < CDate(msk_data_i.Text) Then
        MsgBox "Data final deve ser maior que a data inicial.", vbInformation, "Atenção!"
        msk_data_f.SetFocus
    ElseIf cbo_periodo_i.ListIndex = -1 Then
        MsgBox "Escolha o período inicial.", vbInformation, "Atenção!"
        cbo_periodo_i.SetFocus
    ElseIf cbo_periodo_f.ListIndex = -1 Then
        MsgBox "Escolha o período final.", vbInformation, "Atenção!"
        cbo_periodo_f.SetFocus
    ElseIf cbo_periodo_f.Text < cbo_periodo_i.Text Then
        MsgBox "O periodo final deve ser maior que " & Val(cbo_periodo_i) - 1 & ".", vbInformation, "Atenção!"
        cbo_periodo_f.SetFocus
    ElseIf cboTipoMovimento.ListIndex = -1 Then
        MsgBox "Selecione um tipo de movimento.", vbInformation, "Atenção!"
        cboTipoMovimento.SetFocus
    ElseIf txt_funcionario.Text = "" Then
        MsgBox "Escolha o funcionário.", vbInformation, "Atenção!"
        
        If txt_funcionario.Enabled = True Then
            txt_funcionario.SetFocus
        End If
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
        CalculaDatas
        cbo_periodo_i.ListIndex = 0
        cbo_periodo_f.ListIndex = 3
        
        If cboFuncionario.Enabled = True Then
            cboFuncionario.SetFocus
        End If
       
        If g_nivel_acesso = 4 Then
            PreparaDatas
        End If
    End If
    Screen.MousePointer = 1
End Sub
Private Sub CalculaDatas()
    Dim x_data As Date
    If Day(g_data_def) <= 5 Then
        x_data = CDate("01" & "/" & Month(g_data_def) & "/" & Year(g_data_def)) - 1
        msk_data_f.Text = x_data
        msk_data_i.Text = CDate("01" & "/" & Month(x_data) & "/" & Year(x_data))
    Else
        msk_data_f.Text = g_data_def
        msk_data_i.Text = CDate("01" & "/" & Month(g_data_def) & "/" & Year(g_data_def))
    End If
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
    PreencheCboFuncionario
    PreencheCboPeriodo
    PreencheCboTipoMovimento
    
    If RetiraString(1, gStringChamada) = "Movimento NFCe Auto" Then
'       msk_data_i.Text = fMascaraData(RetiraString(2, gStringChamada))
'       msk_data_f.Text = fMascaraData(RetiraString(2, gStringChamada))
'       cbo_periodo_i.Text = RetiraString(3, gStringChamada)
'       cbo_periodo_f.Text = RetiraString(3, gStringChamada)
    
       If Not PreparaTelaUsuarioLogado(Val(RetiraString(4, gStringChamada))) Then
          gStringChamada = ""
          Unload Me
       End If
       gStringChamada = ""
    Else
       If Not PreparaTelaUsuarioLogado(0) Then
          Me.Hide
       End If
    End If
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_f_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_periodo_i.SetFocus
    End If
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
Private Sub txt_funcionario_GotFocus()
    txt_funcionario.Text = ""
End Sub
Private Sub txt_funcionario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cboFuncionario.Enabled = True Then
            cboFuncionario.SetFocus
        End If
    End If
End Sub
Private Sub txt_funcionario_LostFocus()
    Dim i As Integer
    
    If Val(txt_funcionario.Text) > 0 Then
        If SelecionaFuncionarioNaCombo(Val(txt_funcionario.Text)) Then
            If cmd_imprimir.Enabled = True Then
                cmd_imprimir.SetFocus
            Else
                cmd_sair.SetFocus
            End If
        Else
            If txt_funcionario.Enabled Then
                txt_funcionario.SetFocus
            End If
        End If
    End If
End Sub
