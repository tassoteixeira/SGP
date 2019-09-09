VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form lst_contas_pagar2 
   Caption         =   "Relação de Contas à Pagar Especial"
   ClientHeight    =   4260
   ClientLeft      =   3990
   ClientTop       =   2010
   ClientWidth     =   6795
   Icon            =   "LST_CONTPAG2.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "LST_CONTPAG2.frx":030A
   ScaleHeight     =   4260
   ScaleWidth      =   6795
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   1800
      Picture         =   "LST_CONTPAG2.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Imprime contas à pagar especial (HP)."
      Top             =   3300
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4200
      Picture         =   "LST_CONTPAG2.frx":195A
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   3300
      Width           =   795
   End
   Begin VB.PictureBox Picture1 
      Enabled         =   0   'False
      Height          =   735
      Left            =   2940
      Picture         =   "LST_CONTPAG2.frx":2FEC
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   18
      Top             =   3300
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame frm_dados 
      Height          =   3075
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6555
      Begin VB.ComboBox cboLocalizacaoCobranca 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   2280
         Width           =   4755
      End
      Begin VB.CheckBox chk_rodape 
         Caption         =   "Imprime &Resumo dos Postos"
         Height          =   255
         Left            =   1680
         TabIndex        =   19
         Top             =   2700
         Width           =   2475
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2760
         Picture         =   "LST_CONTPAG2.frx":54EE
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2760
         Picture         =   "LST_CONTPAG2.frx":67C8
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   2760
         Picture         =   "LST_CONTPAG2.frx":7AA2
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   1080
         Width           =   495
      End
      Begin VB.ComboBox cbo_fornecedor 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1920
         Width           =   4755
      End
      Begin VB.ComboBox cbo_conta 
         Height          =   300
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1500
         Width           =   3435
      End
      Begin MSMask.MaskEdBox msk_data_f 
         Height          =   315
         Left            =   1680
         TabIndex        =   8
         Top             =   1080
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
      Begin VB.Label Label6 
         Caption         =   "&Localização Cob."
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   2280
         Width           =   1515
      End
      Begin VB.Label Label4 
         Caption         =   "&Fornecedor"
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   1920
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "&Conta"
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   1500
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
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   1515
      End
      Begin VB.Label Label5 
         Caption         =   "Data de &Emissão"
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1515
      End
   End
   Begin Crystal.CrystalReport cr_relato 
      Left            =   120
      Top             =   3300
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "\VB5\Sgp\Data\lst_baixa_contas_a_pagar_fornecedor.rpt"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "lst_contas_pagar2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_local As Integer
Dim l_linha_tab As Currency
Dim l_coluna_tab As Currency
Dim l_pagina As Integer
Dim l_linha As Currency
Dim l_coluna(0 To 7) As Currency
Dim l_total As Currency
'Dim l_total_1(11, 10) As Currency
'Dim l_total_2(11, 10) As Currency
'Dim l_total_3(11, 10) As Currency
'Dim l_total_4(11, 10) As Currency
'Dim l_total_5(11, 10) As Currency
'Dim l_total_6(11, 10) As Currency
'Dim l_total_7(11, 10) As Currency
'Dim l_total_8(11, 10) As Currency
'Dim l_total_9(11, 10) As Currency
'Dim l_total_10(11, 10) As Currency
'Dim l_total_11(11, 10) As Currency
Dim l_linha_final As Currency

Dim lQtdData As Integer
Dim lQtdLocal As Integer
Dim lQtdEmpresa As Integer
Dim lTotal() As Currency

Dim lRSCriado As Boolean
Dim l_data_hoje As Date
Dim l_data_i As Date
Dim l_data_f As Date
Dim l_nome_conta As String
Dim lCodigoEmpresa(0 To 11) As Integer
Dim lNomeEmpresa(0 To 11) As String
Dim lCodigoGrupoEmpresa As Integer
Dim lNomeGrupoEmpresa As String

Dim lSQL As String
Private rsContaPagar As New ADODB.Recordset

Private Fornecedor As New cFornecedor
Private GrupoEmpresa As New cGrupoEmpresa
Private LocalCobranca As New cLocalCobranca
Private MovimentoContaPagar As New cMovimentoContaPagar
Dim rs As New ADODB.Recordset
Private Sub CalculaResumo()
    Dim i As Integer 'data
    Dim j As Integer 'local de cobranca
    Dim k As Integer 'empresa
    Dim Empresa As Integer
    Dim Coluna As Integer
    Dim xData As Integer
    
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "   SELECT numero_documento, Local_Cobrancas.Nome as NomeLocal"
    lSQL = lSQL & ", valor, data_vencimento, data_emissao"
    lSQL = lSQL & ", Fornecedor.Nome as NomeFornecedor, complemento, local_cobranca, contas_pagar.empresa"
    lSQL = lSQL & "     FROM Contas_Pagar, Local_Cobrancas, Fornecedor"
    lSQL = lSQL & " WHERE Contas_Pagar.Empresa IN ( " & lCodigoEmpresa(1)
    For i = 2 To 11
        If lCodigoEmpresa(i) <> 99 Then
            lSQL = lSQL & ", " & lCodigoEmpresa(i)
        End If
    Next
    lSQL = lSQL & " ) "
    
    'If UCase(lNomeEmpresa(6)) Like "*ANTONIO*" And UCase(g_nome_empresa) Like "*ANTONIO*" Then
    '    lSQL = lSQL & "   WHERE Contas_Pagar.Empresa = " & g_empresa
    '    lSQL = lSQL & " AND"
    'ElseIf UCase(lNomeEmpresa(6)) Like "*ANTONIO*" Then
    '    lSQL = lSQL & "   WHERE Contas_Pagar.Empresa <> " & 6
    '    lSQL = lSQL & " AND"
    'Else
    '    lSQL = lSQL & " WHERE"
    'End If
    
    lSQL = lSQL & "      AND data_vencimento >= " & preparaData(CDate(msk_data_i.Text))
    lSQL = lSQL & "      AND data_vencimento <= " & preparaData(CDate(msk_data_f.Text))
    lSQL = lSQL & "      AND Local_Cobrancas.Codigo = local_cobranca"
    lSQL = lSQL & "      AND Local_Cobrancas.Empresa = contas_pagar.empresa"
    lSQL = lSQL & "      AND Fornecedor.Empresa = contas_pagar.empresa"
    lSQL = lSQL & "      AND Fornecedor.Codigo = codigo_fornecedor"
    lSQL = lSQL & " ORDER BY data_vencimento, local_cobranca, contas_pagar.empresa"
    
    
    'If UCase(lNomeEmpresa(6)) Like "*ANTONIO*" And UCase(g_nome_empresa) Like "*ANTONIO*" Then
    '    For i = 1 To 5
    '        lNomeEmpresa(i) = ""
    '    Next
    'ElseIf UCase(lNomeEmpresa(6)) Like "*ANTONIO*" Then
    '    lNomeEmpresa(6) = ""
    'End If
    
    
    'Abre RecordSet
    Set rsContaPagar = New ADODB.Recordset
    Set rsContaPagar = Conectar.RsConexao(lSQL)
    
    
    CriaRS
    GravaRS
    rs.Sort = "data_vencimento, local_cobranca, empresa"
    
    
    'Verifica movimento
    If rs.RecordCount > 0 Then
        Do Until rs.EOF
            Empresa = 0
            Coluna = 0
            For i = 1 To 11
                If lCodigoEmpresa(i) = rs("empresa").Value Then
                    Empresa = lCodigoEmpresa(i)
                    Coluna = i
                End If
            Next
        
            xData = DateDiff("d", CDate(msk_data_i.Text), rs("data_vencimento").Value) + 1
            For i = 0 To xData
                For j = 0 To rs("local_cobranca").Value
                    For k = 0 To Coluna
                        lTotal(i, j, k) = lTotal(i, j, k) + rs("valor").Value
                        If k = 0 Then
                            k = Coluna - 1
                        End If
                    Next
                    If j = 0 Then
                        j = rs("local_cobranca").Value - 1
                    End If
                Next
                If i = 0 Then
                    i = xData - 1
                End If
            Next
            'For i = 0 To xData
            '    For j = 0 To rs("local_cobranca").Value
            '        For k = 0 To rs("empresa").Value
            '            lTotal(i, j, k) = lTotal(i, j, k) + rs("valor").Value
            '            If k = 0 Then
            '                k = rs("empresa").Value - 1
            '            End If
            '        Next
            '        If j = 0 Then
            '            j = rs("local_cobranca").Value - 1
            '        End If
            '    Next
            '    If i = 0 Then
            '        i = xData - 1
            '    End If
            'Next
            rs.MoveNext
        Loop
    End If
    If rsContaPagar.State = 1 Then
        rsContaPagar.Close
        rs.Close
    End If
End Sub
Private Sub CriaRS()
    With rs
        'If lRSCriado Then
        '    .MoveFirst
        '    Do Until .EOF
        '        .Delete
        '        .MoveNext
        '    Loop
        'Else
            .CursorLocation = adUseClient
            .Fields.Append "numero_documento", adVarChar, 10
            .Fields.Append "NomeLocal", adVarChar, 30
            .Fields.Append "valor", adCurrency
            .Fields.Append "data_vencimento", adDate
            .Fields.Append "data_emissao", adDate
            .Fields.Append "NomeFornecedor", adVarChar, 40
            .Fields.Append "complemento", adVarChar, 40
            .Fields.Append "local_cobranca", adInteger
            .Fields.Append "empresa", adInteger
            .Open
            lRSCriado = True
        'End If
    End With
End Sub
Private Sub GravaRS()
    If rsContaPagar.RecordCount > 0 Then
        rsContaPagar.MoveFirst
        Do Until rsContaPagar.EOF
            rs.AddNew
            rs!Numero_Documento = rsContaPagar("numero_documento").Value
            rs!NomeLocal = rsContaPagar("NomeLocal").Value
            rs!Valor = rsContaPagar("valor").Value
            rs!Data_Vencimento = rsContaPagar("data_vencimento").Value
            rs!Data_Emissao = rsContaPagar("data_emissao").Value
            rs!NomeFornecedor = rsContaPagar("nomefornecedor").Value
            rs!Complemento = rsContaPagar("Complemento").Value
            rs!local_cobranca = rsContaPagar("local_cobranca").Value
            rs!Empresa = rsContaPagar("Empresa").Value
            rs.Update
            rsContaPagar.MoveNext
        Loop
    End If
End Sub
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set rsContaPagar = Nothing
    Set Fornecedor = Nothing
    Set GrupoEmpresa = Nothing
    Set LocalCobranca = Nothing
    Set MovimentoContaPagar = Nothing
End Sub
Private Sub PreencheCboConta()
    Dim rstConta As New ADODB.Recordset
    
    cbo_conta.Clear
    cbo_conta.AddItem "Todas as Contas"
    cbo_conta.ItemData(cbo_conta.NewIndex) = 0
    Set rstConta = Conectar.RsConexao("SELECT Codigo, Nome FROM Contas WHERE Empresa = " & g_empresa & " ORDER BY Nome")
    'loop RecordSet
    With rstConta
        If Not .BOF Or Not .EOF Then
            .MoveFirst
            Do Until .EOF
                cbo_conta.AddItem !Nome
                cbo_conta.ItemData(cbo_conta.NewIndex) = !Codigo
                .MoveNext
            Loop
        End If
        .Close
    End With
    Set rstConta = Nothing
End Sub
Private Sub Relatorio()
    Dim i As Integer
    Dim i2 As Integer
    ZeraVariaveis
    l_data_hoje = CDate(msk_data.Text)
    l_data_i = msk_data_i.Text
    l_data_f = msk_data_f.Text
    l_nome_conta = cbo_conta.Text
    l_linha = -0.7
    l_pagina = 0
    l_total = 0
    l_coluna(0) = 0
    l_coluna(1) = 1.5
    l_coluna(2) = 5
    l_coluna(3) = 6.7
    l_coluna(4) = 8.4
    l_coluna(5) = 9.9
    l_coluna(6) = 16.25
    l_coluna(7) = 20
    
    
    If chk_rodape.Value = 1 Then
        CalculaResumo
        ImpResumo
    Else
        'Prepara SQL
        lSQL = ""
        lSQL = lSQL & "   SELECT numero_documento, Local_Cobrancas.Nome as NomeLocal"
        lSQL = lSQL & ", valor, data_vencimento, data_emissao"
        lSQL = lSQL & ", Fornecedor.Nome as NomeFornecedor, complemento"
        lSQL = lSQL & "     FROM Contas_Pagar, Local_Cobrancas, Fornecedor"
        lSQL = lSQL & "    WHERE contas_pagar.empresa = " & g_empresa
        lSQL = lSQL & "      AND data_vencimento >= " & preparaData(CDate(msk_data_i.Text))
        lSQL = lSQL & "      AND data_vencimento <= " & preparaData(CDate(msk_data_f.Text))
        lSQL = lSQL & "      AND Local_Cobrancas.Codigo = local_cobranca"
        lSQL = lSQL & "      AND Local_Cobrancas.Empresa = contas_pagar.empresa"
        lSQL = lSQL & "      AND Fornecedor.Empresa = contas_pagar.empresa"
        lSQL = lSQL & "      AND Fornecedor.Codigo = codigo_fornecedor"
        If cbo_fornecedor.ItemData(cbo_fornecedor.ListIndex) > 0 Then
            lSQL = lSQL & "      AND codigo_fornecedor = " & Val(cbo_fornecedor.ItemData(cbo_fornecedor.ListIndex))
        End If
        If cboLocalizacaoCobranca.ItemData(cboLocalizacaoCobranca.ListIndex) > 0 Then
            lSQL = lSQL & "      AND local_cobranca = " & Val(cboLocalizacaoCobranca.ItemData(cboLocalizacaoCobranca.ListIndex))
        End If
        lSQL = lSQL & " ORDER BY data_vencimento, data_emissao"
        
        'Abre RecordSet
        Set rsContaPagar = New ADODB.Recordset
        Set rsContaPagar = Conectar.RsConexao(lSQL)
        
        'Verifica movimento
        If rsContaPagar.RecordCount > 0 Then
            ImpDados
        End If
        If rsContaPagar.State = 1 Then
            rsContaPagar.Close
        End If
    End If
    cmd_sair.SetFocus
End Sub
Private Sub ImpDados()
    'seleciona medidas para centímetros
    Printer.ScaleMode = 7
    Printer.FontName = "Times New Roman"
    Printer.FontName = "Times New Roman"
    Printer.FontSize = 10
    Printer.FontBold = False
    'loop movimento de contas à pagar
    Do Until rsContaPagar.EOF
        If l_linha >= 50 Then
            ImpGrade l_linha
            Printer.NewPage
            l_linha = -0.7
        End If
        'calcula linha e coluna
        l_linha = l_linha + 1
        l_linha_tab = l_linha * 0.4 + 1 + 3 + 0.5
        l_coluna_tab = 0
        'move dados
        'imprime detalhe
        If l_linha_tab = 4.62 Then
            Printer.FontSize = 8
            ImprimeTexto " ", l_coluna(0), l_coluna(1), l_linha_tab, l_local
        End If
        Printer.FontSize = 8
        ImprimeTexto rsContaPagar("numero_documento").Value, l_coluna(0), l_coluna(1), l_linha_tab, l_local
        ImprimeTexto rsContaPagar("NomeLocal").Value, l_coluna(1), l_coluna(2), l_linha_tab, l_local
        ImprimeValor Format(rsContaPagar("valor").Value, "##,###,##0.00"), l_coluna(2), l_coluna(3), l_linha_tab, l_local
        ImprimeCentralizado Format(rsContaPagar("data_vencimento").Value, "dd/mm/yyyy"), l_coluna(3), l_coluna(4), l_linha_tab, l_local
        ImprimeCentralizado Format(rsContaPagar("data_emissao").Value, "dd/mm/yyyy"), l_coluna(4), l_coluna(5), l_linha_tab, l_local
        ImprimeTexto rsContaPagar("NomeFornecedor").Value, l_coluna(5), l_coluna(6), l_linha_tab, l_local
        ImprimeTexto rsContaPagar("complemento").Value, l_coluna(6), l_coluna(7), l_linha_tab, l_local
        l_total = l_total + rsContaPagar("valor").Value
        rsContaPagar.MoveNext
    Loop
    ImpGrade l_linha
    ImpTotal
    Printer.EndDoc
End Sub
Private Sub ImpTotal()
    If l_pagina > 0 Then
        l_linha_final = l_linha * 0.4 + l_linha_tab + 3 + 0.5 + 0.5 + 0.8
        Printer.DrawWidth = 2
        Printer.Line (l_coluna_tab + l_coluna(0), l_linha_final - 0.8)-(l_coluna_tab + l_coluna(0), l_linha_final)
        Printer.Line (l_coluna_tab + l_coluna(2), l_linha_final - 0.8)-(l_coluna_tab + l_coluna(2), l_linha_final)
        Printer.Line (l_coluna_tab + l_coluna(3), l_linha_final - 0.8)-(l_coluna_tab + l_coluna(3), l_linha_final)
        Printer.Line (l_coluna_tab + l_coluna(7), l_linha_final - 0.8)-(l_coluna_tab + l_coluna(7), l_linha_final)
        Printer.Line (l_coluna_tab + l_coluna(0), l_linha_final)-(l_coluna_tab + l_coluna(7), l_linha_final)
        Printer.DrawWidth = 4
        Printer.FontBold = True
        l_linha = l_linha + 1
        l_linha_tab = l_linha * 0.4 + 1 + 3 + 0.5 + 0.3
        ImprimeValor "Total.:", l_coluna(0), l_coluna(2), l_linha_tab, l_local
        ImprimeValor Format(l_total, "##,###,##0.00"), l_coluna(2), l_coluna(3), l_linha_tab, l_local
        Printer.FontBold = False
    End If
End Sub
Private Sub ImpResumoCab(ByVal xUltimaLinha As Currency)
    Dim i As Integer 'data
    Dim j As Integer 'local de cobranca
    Dim k As Integer 'empresa
    Dim xData As Date
    Dim xLinhaFinal As Currency
    
    On Error GoTo ErroImpResumoCab
    
    'seleciona medidas para centímetros
    Printer.ScaleMode = 7
    If l_pagina = 0 Then
        Printer.PaperSize = vbPRPSLetter 'vbPRPSLegal
        Printer.Orientation = vbPRORLandscape
    End If
    Printer.FontName = "Times New Roman"
    Printer.FontName = "Times New Roman"
    Printer.FontSize = 10
    Printer.FontBold = False
    
    'calcula linha e coluna
    l_linha = 1
    l_linha_tab = 1
    l_coluna_tab = 0.5
    'move dados
    l_coluna(7) = 26
    If l_pagina = 0 Then
        'imprime detalhe
        Printer.FontSize = 8
        ImprimeTexto " ", l_coluna(0), l_coluna(1), l_linha_tab, l_local
    Else
        Printer.Line (l_coluna_tab + 0, l_linha_tab + 3)-(l_coluna_tab + 0, xUltimaLinha + 0.4)
        Printer.Line (l_coluna_tab + 2, l_linha_tab + 3)-(l_coluna_tab + 2, xUltimaLinha + 0.4)
        Printer.Line (l_coluna_tab + 4, l_linha_tab + 3)-(l_coluna_tab + 4, xUltimaLinha + 0.4)
        Printer.Line (l_coluna_tab + 6, l_linha_tab + 3)-(l_coluna_tab + 6, xUltimaLinha + 0.4)
        Printer.Line (l_coluna_tab + 8, l_linha_tab + 3)-(l_coluna_tab + 8, xUltimaLinha + 0.4)
        Printer.Line (l_coluna_tab + 10, l_linha_tab + 3)-(l_coluna_tab + 10, xUltimaLinha + 0.4)
        Printer.Line (l_coluna_tab + 12, l_linha_tab + 3)-(l_coluna_tab + 12, xUltimaLinha + 0.4)
        Printer.Line (l_coluna_tab + 14, l_linha_tab + 3)-(l_coluna_tab + 14, xUltimaLinha + 0.4)
        Printer.Line (l_coluna_tab + 16, l_linha_tab + 3)-(l_coluna_tab + 16, xUltimaLinha + 0.4)
        Printer.Line (l_coluna_tab + 18, l_linha_tab + 3)-(l_coluna_tab + 18, xUltimaLinha + 0.4)
        Printer.Line (l_coluna_tab + 20, l_linha_tab + 3)-(l_coluna_tab + 20, xUltimaLinha + 0.4)
        Printer.Line (l_coluna_tab + 22, l_linha_tab + 3)-(l_coluna_tab + 22, xUltimaLinha + 0.4)
        Printer.Line (l_coluna_tab + 24, l_linha_tab + 3)-(l_coluna_tab + 24, xUltimaLinha + 0.4)
        Printer.Line (l_coluna_tab + 26, l_linha_tab + 3)-(l_coluna_tab + 26, xUltimaLinha + 0.4)
        Printer.NewPage
    End If
    
    Printer.FontSize = 15
    Printer.CurrentX = l_coluna_tab + 0.5
    Printer.CurrentY = l_linha_tab + 0.1
    Printer.Print lNomeGrupoEmpresa
    Printer.CurrentX = l_coluna_tab + 0.5
    Printer.CurrentY = l_linha_tab + 0.7
    Printer.Print "Resumo do Contas à Pagar"
    
    Printer.FontSize = 10
    Printer.CurrentX = l_coluna_tab + 23.5
    Printer.CurrentY = l_linha_tab + 0.16
    l_pagina = l_pagina + 1
    Printer.Print "Página,   " & Format(l_pagina, "000")
    Printer.CurrentX = l_coluna_tab + 19.5
    Printer.CurrentY = l_linha_tab + 0.7
    Printer.Print Trim(g_cidade_empresa) & ", " & Format(msk_data.Text, "dd/mm/yyyy")
    
    Printer.PaintPicture Picture1, 25, 2.3
    Printer.CurrentX = l_coluna_tab + 0.5
    Printer.CurrentY = l_linha_tab + 1.4
    Printer.Print "Período de Referência.: " & Format(msk_data_i.Text, "dd/mm/yyyy") & " a "; Format(msk_data_f.Text, "dd/mm/yyyy")
    Printer.CurrentX = l_coluna_tab + 0.5
    Printer.CurrentY = l_linha_tab + 1.9
    Printer.Print "Conta.........................: " & l_nome_conta
    Printer.CurrentX = l_coluna_tab + 0.5
    Printer.CurrentY = l_linha_tab + 2.4
    Printer.Print "Fornecedor.................: " & l_nome_conta
    
    Printer.DrawWidth = 2
    'Linhas Horizontais
    Printer.Line (l_coluna_tab + l_coluna(0), l_linha_tab)-(l_coluna_tab + l_coluna(7), l_linha_tab)
    Printer.Line (l_coluna_tab + l_coluna(0), l_linha_tab + 3)-(l_coluna_tab + l_coluna(7), l_linha_tab + 3)
    Printer.Line (l_coluna_tab + l_coluna(0), l_linha_tab + 3 + 0.5)-(l_coluna_tab + l_coluna(7), l_linha_tab + 3 + 0.5)
    'Linhas Verticais
    Printer.Line (l_coluna_tab + l_coluna(0), l_linha_tab)-(l_coluna_tab + l_coluna(0), l_linha_tab + 3 + 0.5)
    Printer.Line (l_coluna_tab + l_coluna(7), l_linha_tab)-(l_coluna_tab + l_coluna(7), l_linha_tab + 3 + 0.5)
    
    l_linha_final = l_linha_tab + 3
    ImprimeCentralizado "Venc.", l_coluna_tab + 0, l_coluna_tab + 2, l_linha_final + 0.05, l_local
    ImprimeCentralizado Trim(lNomeEmpresa(1)), l_coluna_tab + 2, l_coluna_tab + 4, l_linha_final + 0.05, l_local
    ImprimeCentralizado Trim(lNomeEmpresa(2)), l_coluna_tab + 4, l_coluna_tab + 6, l_linha_final + 0.05, l_local
    ImprimeCentralizado Trim(lNomeEmpresa(3)), l_coluna_tab + 6, l_coluna_tab + 8, l_linha_final + 0.05, l_local
    ImprimeCentralizado Trim(lNomeEmpresa(4)), l_coluna_tab + 8, l_coluna_tab + 10, l_linha_final + 0.05, l_local
    ImprimeCentralizado Trim(lNomeEmpresa(5)), l_coluna_tab + 10, l_coluna_tab + 12, l_linha_final + 0.05, l_local
    ImprimeCentralizado Trim(lNomeEmpresa(6)), l_coluna_tab + 12, l_coluna_tab + 14, l_linha_final + 0.05, l_local
    ImprimeCentralizado Trim(lNomeEmpresa(7)), l_coluna_tab + 14, l_coluna_tab + 16, l_linha_final + 0.05, l_local
    ImprimeCentralizado Trim(lNomeEmpresa(8)), l_coluna_tab + 16, l_coluna_tab + 18, l_linha_final + 0.05, l_local
    ImprimeCentralizado Trim(lNomeEmpresa(9)), l_coluna_tab + 18, l_coluna_tab + 20, l_linha_final + 0.05, l_local
    ImprimeCentralizado Trim(lNomeEmpresa(10)), l_coluna_tab + 20, l_coluna_tab + 22, l_linha_final + 0.05, l_local
    ImprimeCentralizado Trim(lNomeEmpresa(11)), l_coluna_tab + 22, l_coluna_tab + 24, l_linha_final + 0.05, l_local
    ImprimeCentralizado "Total", 24, 26, l_linha_final + 0.05, l_local
    
    Printer.DrawWidth = 4
    Printer.FontBold = True
    ''ImprimeValor "Total.:", 0, 2, l_linha_final + 4.9, l_local
    l_linha_final = l_linha_final + 0.1
    Exit Sub

ErroImpResumoCab:
    Resume Next
End Sub
Private Sub ImpResumo()
    Dim i As Integer 'data
    Dim j As Integer 'local de cobranca
    Dim k As Integer 'empresa
    Dim xData As Date
    Dim xLinhaFinal As Currency
    
    If l_pagina = 0 Then
        Call ImpResumoCab(0)
    End If
    
    For i = 1 To lQtdData
        'testa se o total da data "i" é maior que zero
        If lTotal(i, 0, 0) > 0 Then
            For j = 1 To lQtdLocal
                'testa se o total do local de cobranca "j" é maior que zero
                'Imprime total do local de cobrança da data
                If lTotal(i, j, 0) > 0 Then
                    If l_linha_final >= 18 Then
                        Call ImpResumoCab(l_linha_final)
                    End If
                    l_linha_final = l_linha_final + 0.5
                    If LocalCobranca.LocalizarCodigo(j, k) Then
                        ImprimeTexto LocalCobranca.Nome, l_coluna_tab + 0, l_coluna_tab + 2, l_linha_final, l_local
                    End If
                    ImprimeValor Format(lTotal(i, j, 1), "##,###,##0.00"), l_coluna_tab + 2, l_coluna_tab + 4, l_linha_final, l_local
                    ImprimeValor Format(lTotal(i, j, 2), "##,###,##0.00"), l_coluna_tab + 4, l_coluna_tab + 6, l_linha_final, l_local
                    ImprimeValor Format(lTotal(i, j, 3), "##,###,##0.00"), l_coluna_tab + 6, l_coluna_tab + 8, l_linha_final, l_local
                    ImprimeValor Format(lTotal(i, j, 4), "##,###,##0.00"), l_coluna_tab + 8, l_coluna_tab + 10, l_linha_final, l_local
                    ImprimeValor Format(lTotal(i, j, 5), "##,###,##0.00"), l_coluna_tab + 10, l_coluna_tab + 12, l_linha_final, l_local
                    ImprimeValor Format(lTotal(i, j, 6), "##,###,##0.00"), l_coluna_tab + 12, l_coluna_tab + 14, l_linha_final, l_local
                    ImprimeValor Format(lTotal(i, j, 7), "##,###,##0.00"), l_coluna_tab + 14, l_coluna_tab + 16, l_linha_final, l_local
                    ImprimeValor Format(lTotal(i, j, 8), "##,###,##0.00"), l_coluna_tab + 16, l_coluna_tab + 18, l_linha_final, l_local
                    ImprimeValor Format(lTotal(i, j, 9), "##,###,##0.00"), l_coluna_tab + 18, l_coluna_tab + 20, l_linha_final, l_local
                    ImprimeValor Format(lTotal(i, j, 10), "##,###,##0.00"), l_coluna_tab + 20, l_coluna_tab + 22, l_linha_final, l_local
                    ImprimeValor Format(lTotal(i, j, 11), "##,###,##0.00"), l_coluna_tab + 22, l_coluna_tab + 24, l_linha_final, l_local
                    ImprimeValor Format(lTotal(i, j, 0), "##,###,##0.00"), l_coluna_tab + 24, l_coluna_tab + 26, l_linha_final, l_local
                End If
            Next
            'Imprime total da data
            xData = CDate(msk_data_i.Text) + i - 1
            l_linha_final = l_linha_final + 0.5
            ImprimeCentralizado Format(xData, "dd/mm/yyyy"), l_coluna_tab + 0, l_coluna_tab + 2, l_linha_final, l_local
            ImprimeValor Format(lTotal(i, 0, 1), "##,###,##0.00"), l_coluna_tab + 2, l_coluna_tab + 4, l_linha_final, l_local
            ImprimeValor Format(lTotal(i, 0, 2), "##,###,##0.00"), l_coluna_tab + 4, l_coluna_tab + 6, l_linha_final, l_local
            ImprimeValor Format(lTotal(i, 0, 3), "##,###,##0.00"), l_coluna_tab + 6, l_coluna_tab + 8, l_linha_final, l_local
            ImprimeValor Format(lTotal(i, 0, 4), "##,###,##0.00"), l_coluna_tab + 8, l_coluna_tab + 10, l_linha_final, l_local
            ImprimeValor Format(lTotal(i, 0, 5), "##,###,##0.00"), l_coluna_tab + 10, l_coluna_tab + 12, l_linha_final, l_local
            ImprimeValor Format(lTotal(i, 0, 6), "##,###,##0.00"), l_coluna_tab + 12, l_coluna_tab + 14, l_linha_final, l_local
            ImprimeValor Format(lTotal(i, 0, 7), "##,###,##0.00"), l_coluna_tab + 14, l_coluna_tab + 16, l_linha_final, l_local
            ImprimeValor Format(lTotal(i, 0, 8), "##,###,##0.00"), l_coluna_tab + 16, l_coluna_tab + 18, l_linha_final, l_local
            ImprimeValor Format(lTotal(i, 0, 9), "##,###,##0.00"), l_coluna_tab + 18, l_coluna_tab + 20, l_linha_final, l_local
            ImprimeValor Format(lTotal(i, 0, 10), "##,###,##0.00"), l_coluna_tab + 20, l_coluna_tab + 22, l_linha_final, l_local
            ImprimeValor Format(lTotal(i, 0, 11), "##,###,##0.00"), l_coluna_tab + 22, l_coluna_tab + 24, l_linha_final, l_local
            ImprimeValor Format(lTotal(i, 0, 0), "##,###,##0.00"), l_coluna_tab + 24, l_coluna_tab + 26, l_linha_final, l_local
            Printer.Line (l_coluna_tab + l_coluna(0), l_linha_final + 0.5)-(l_coluna_tab + l_coluna(7), l_linha_final + 0.5)
            l_linha_final = l_linha_final + 0.1
        End If
    Next
    
    

    'Total Geral
    For j = 1 To lQtdLocal
        'testa se o total do local de cobranca "j" é maior que zero
        'Imprime total do local de cobrança da data
        If lTotal(0, j, 0) > 0 Then
            If l_linha_final > 18 Then
                Call ImpResumoCab(l_linha_final)
            End If
            l_linha_final = l_linha_final + 0.5
            If LocalCobranca.LocalizarCodigo(j, k) Then
                ImprimeTexto LocalCobranca.Nome, l_coluna_tab + 0, l_coluna_tab + 2, l_linha_final, l_local
            End If
            ImprimeValor Format(lTotal(0, j, 1), "##,###,##0.00"), l_coluna_tab + 2, l_coluna_tab + 4, l_linha_final, l_local
            ImprimeValor Format(lTotal(0, j, 2), "##,###,##0.00"), l_coluna_tab + 4, l_coluna_tab + 6, l_linha_final, l_local
            ImprimeValor Format(lTotal(0, j, 3), "##,###,##0.00"), l_coluna_tab + 6, l_coluna_tab + 8, l_linha_final, l_local
            ImprimeValor Format(lTotal(0, j, 4), "##,###,##0.00"), l_coluna_tab + 8, l_coluna_tab + 10, l_linha_final, l_local
            ImprimeValor Format(lTotal(0, j, 5), "##,###,##0.00"), l_coluna_tab + 10, l_coluna_tab + 12, l_linha_final, l_local
            ImprimeValor Format(lTotal(0, j, 6), "##,###,##0.00"), l_coluna_tab + 12, l_coluna_tab + 14, l_linha_final, l_local
            ImprimeValor Format(lTotal(0, j, 7), "##,###,##0.00"), l_coluna_tab + 14, l_coluna_tab + 16, l_linha_final, l_local
            ImprimeValor Format(lTotal(0, j, 8), "##,###,##0.00"), l_coluna_tab + 16, l_coluna_tab + 18, l_linha_final, l_local
            ImprimeValor Format(lTotal(0, j, 9), "##,###,##0.00"), l_coluna_tab + 18, l_coluna_tab + 20, l_linha_final, l_local
            ImprimeValor Format(lTotal(0, j, 10), "##,###,##0.00"), l_coluna_tab + 20, l_coluna_tab + 22, l_linha_final, l_local
            ImprimeValor Format(lTotal(0, j, 11), "##,###,##0.00"), l_coluna_tab + 22, l_coluna_tab + 24, l_linha_final, l_local
            ImprimeValor Format(lTotal(0, j, 0), "##,###,##0.00"), l_coluna_tab + 24, l_coluna_tab + 26, l_linha_final, l_local
        End If
    Next
    'Imprime Total Geral
    l_linha_final = l_linha_final + 0.6
    ImprimeCentralizado "TOTAL", l_coluna_tab + 0, l_coluna_tab + 2, l_linha_final, l_local
    ImprimeValor Format(lTotal(0, 0, 1), "##,###,##0.00"), l_coluna_tab + 2, l_coluna_tab + 4, l_linha_final, l_local
    ImprimeValor Format(lTotal(0, 0, 2), "##,###,##0.00"), l_coluna_tab + 4, l_coluna_tab + 6, l_linha_final, l_local
    ImprimeValor Format(lTotal(0, 0, 3), "##,###,##0.00"), l_coluna_tab + 6, l_coluna_tab + 8, l_linha_final, l_local
    ImprimeValor Format(lTotal(0, 0, 4), "##,###,##0.00"), l_coluna_tab + 8, l_coluna_tab + 10, l_linha_final, l_local
    ImprimeValor Format(lTotal(0, 0, 5), "##,###,##0.00"), l_coluna_tab + 10, l_coluna_tab + 12, l_linha_final, l_local
    ImprimeValor Format(lTotal(0, 0, 6), "##,###,##0.00"), l_coluna_tab + 12, l_coluna_tab + 14, l_linha_final, l_local
    ImprimeValor Format(lTotal(0, 0, 7), "##,###,##0.00"), l_coluna_tab + 14, l_coluna_tab + 16, l_linha_final, l_local
    ImprimeValor Format(lTotal(0, 0, 8), "##,###,##0.00"), l_coluna_tab + 16, l_coluna_tab + 18, l_linha_final, l_local
    ImprimeValor Format(lTotal(0, 0, 9), "##,###,##0.00"), l_coluna_tab + 18, l_coluna_tab + 20, l_linha_final, l_local
    ImprimeValor Format(lTotal(0, 0, 10), "##,###,##0.00"), l_coluna_tab + 20, l_coluna_tab + 22, l_linha_final, l_local
    ImprimeValor Format(lTotal(0, 0, 11), "##,###,##0.00"), l_coluna_tab + 22, l_coluna_tab + 24, l_linha_final, l_local
    ImprimeValor Format(lTotal(0, 0, 0), "##,###,##0.00"), l_coluna_tab + 24, l_coluna_tab + 26, l_linha_final, l_local
    Printer.Line (l_coluna_tab + l_coluna(0), l_linha_final + 0.5)-(l_coluna_tab + l_coluna(7), l_linha_final + 0.5)
    
    
    xLinhaFinal = l_linha_final
    l_linha_final = l_linha_tab + 3
    Printer.DrawWidth = 2
    Printer.Line (l_coluna_tab + l_coluna(0), l_linha_final)-(l_coluna_tab + l_coluna(7), l_linha_final)
    Printer.Line (l_coluna_tab + l_coluna(0), l_linha_final + 0.5)-(l_coluna_tab + l_coluna(7), l_linha_final + 0.5)
    Printer.Line (l_coluna_tab + 0, l_linha_tab)-(l_coluna_tab + 0, 1.5 + xLinhaFinal - 1)
    Printer.Line (l_coluna_tab + 2, l_linha_final)-(l_coluna_tab + 2, 1.5 + xLinhaFinal - 1)
    Printer.Line (l_coluna_tab + 4, l_linha_final)-(l_coluna_tab + 4, 1.5 + xLinhaFinal - 1)
    Printer.Line (l_coluna_tab + 6, l_linha_final)-(l_coluna_tab + 6, 1.5 + xLinhaFinal - 1)
    Printer.Line (l_coluna_tab + 8, l_linha_final)-(l_coluna_tab + 8, 1.5 + xLinhaFinal - 1)
    Printer.Line (l_coluna_tab + 10, l_linha_final)-(l_coluna_tab + 10, 1.5 + xLinhaFinal - 1)
    Printer.Line (l_coluna_tab + 12, l_linha_final)-(l_coluna_tab + 12, 1.5 + xLinhaFinal - 1)
    Printer.Line (l_coluna_tab + 14, l_linha_final)-(l_coluna_tab + 14, 1.5 + xLinhaFinal - 1)
    Printer.Line (l_coluna_tab + 16, l_linha_final)-(l_coluna_tab + 16, 1.5 + xLinhaFinal - 1)
    Printer.Line (l_coluna_tab + 18, l_linha_final)-(l_coluna_tab + 18, 1.5 + xLinhaFinal - 1)
    Printer.Line (l_coluna_tab + 20, l_linha_final)-(l_coluna_tab + 20, 1.5 + xLinhaFinal - 1)
    Printer.Line (l_coluna_tab + 22, l_linha_final)-(l_coluna_tab + 22, 1.5 + xLinhaFinal - 1)
    Printer.Line (l_coluna_tab + 24, l_linha_final)-(l_coluna_tab + 24, 1.5 + xLinhaFinal - 1)
    Printer.Line (l_coluna_tab + 26, l_linha_tab)-(l_coluna_tab + 26, 1.5 + xLinhaFinal - 1)
    Printer.Line (l_coluna_tab + l_coluna(0), xLinhaFinal)-(l_coluna_tab + l_coluna(7), xLinhaFinal)
    Printer.Line (l_coluna_tab + l_coluna(0), xLinhaFinal + 0.5)-(l_coluna_tab + l_coluna(7), xLinhaFinal + 0.5)
    
    Printer.FontBold = False
    Printer.EndDoc
End Sub
Private Sub ImpGrade(x_vezes As Currency)
    Dim i As Integer
    l_pagina = l_pagina + 1
    l_coluna_tab = 0
    l_linha_tab = 1
    
    Printer.FontSize = 15
    Printer.CurrentX = l_coluna_tab + 0.5
    Printer.CurrentY = l_linha_tab + 0.1
    Printer.Print g_nome_empresa
    Printer.CurrentX = l_coluna_tab + 0.5
    Printer.CurrentY = l_linha_tab + 0.7
    Printer.Print "Relação dos Vencimentos"
    
    Printer.FontSize = 10
    Printer.CurrentX = l_coluna_tab + 16.5
    Printer.CurrentY = l_linha_tab + 0.16
    Printer.Print "Página,   " & Format(l_pagina, "000")
    Printer.CurrentX = l_coluna_tab + 16.5
    Printer.CurrentY = l_linha_tab + 0.7
    Printer.Print "Goiânia, " & Format(msk_data.Text, "dd/mm/yyyy")
    
    Printer.PaintPicture Picture1, 18, 2.3
    Printer.CurrentX = l_coluna_tab + 0.5
    Printer.CurrentY = l_linha_tab + 1.4
    Printer.Print "Período de Referência.: " & Format(msk_data_i.Text, "dd/mm/yyyy") & " a "; Format(msk_data_f.Text, "dd/mm/yyyy")
    Printer.CurrentX = l_coluna_tab + 0.5
    Printer.CurrentY = l_linha_tab + 1.9
    Printer.Print "Conta.........................: " & l_nome_conta
    Printer.CurrentX = l_coluna_tab + 0.5
    Printer.CurrentY = l_linha_tab + 2.4
    Printer.Print "Fornecedor.................: " & l_nome_conta
    
    Printer.DrawWidth = 2
    Printer.Line (l_coluna_tab + l_coluna(0), l_linha_tab)-(l_coluna_tab + l_coluna(7), l_linha_tab)
    Printer.Line (l_coluna_tab + l_coluna(0), l_linha_tab + 3)-(l_coluna_tab + l_coluna(7), l_linha_tab + 3)
    Printer.Line (l_coluna_tab + l_coluna(0), l_linha_tab + 3 + 0.5)-(l_coluna_tab + l_coluna(7), l_linha_tab + 3 + 0.5)
    Printer.Line (l_coluna_tab + l_coluna(0), l_linha * 0.4 + 1 + 3 + 0.5 + 0.5)-(l_coluna_tab + l_coluna(7), l_linha * 0.4 + 1 + 3 + 0.5 + 0.5)
    
    Printer.Line (l_coluna_tab + l_coluna(0), l_linha_tab)-(l_coluna_tab + l_coluna(0), l_linha * 0.4 + 1 + 3 + 0.5 + 0.5)
    Printer.Line (l_coluna_tab + l_coluna(1), l_linha_tab + 3)-(l_coluna_tab + l_coluna(1), l_linha * 0.4 + 1 + 3 + 0.5 + 0.5)
    Printer.Line (l_coluna_tab + l_coluna(2), l_linha_tab + 3)-(l_coluna_tab + l_coluna(2), l_linha * 0.4 + 1 + 3 + 0.5 + 0.5)
    Printer.Line (l_coluna_tab + l_coluna(3), l_linha_tab + 3)-(l_coluna_tab + l_coluna(3), l_linha * 0.4 + 1 + 3 + 0.5 + 0.5)
    Printer.Line (l_coluna_tab + l_coluna(4), l_linha_tab + 3)-(l_coluna_tab + l_coluna(4), l_linha * 0.4 + 1 + 3 + 0.5 + 0.5)
    Printer.Line (l_coluna_tab + l_coluna(5), l_linha_tab + 3)-(l_coluna_tab + l_coluna(5), l_linha * 0.4 + 1 + 3 + 0.5 + 0.5)
    Printer.Line (l_coluna_tab + l_coluna(6), l_linha_tab + 3)-(l_coluna_tab + l_coluna(6), l_linha * 0.4 + 1 + 3 + 0.5 + 0.5)
    Printer.Line (l_coluna_tab + l_coluna(7), l_linha_tab)-(l_coluna_tab + l_coluna(7), l_linha * 0.4 + 1 + 3 + 0.5 + 0.5)
    
    Printer.FontSize = 10
    Printer.DrawWidth = 4
    l_linha_tab = l_linha_tab + 3 + 0.1
    ImprimeTexto "N.Doc.", l_coluna(0), l_coluna(1), l_linha_tab, l_local
    ImprimeTexto "Local", l_coluna(1), l_coluna(2), l_linha_tab, l_local
    ImprimeCentralizado "Valor", l_coluna(2), l_coluna(3), l_linha_tab, l_local
    ImprimeCentralizado "Vencim.", l_coluna(3), l_coluna(4), l_linha_tab, l_local
    ImprimeCentralizado "Emissão", l_coluna(4), l_coluna(5), l_linha_tab, l_local
    ImprimeTexto "Fornecedor", l_coluna(5), l_coluna(6), l_linha_tab, l_local
    ImprimeTexto "Complemento", l_coluna(6), l_coluna(7), l_linha_tab, l_local
    l_coluna_tab = 0
    l_linha_tab = 1
End Sub
Private Sub ZeraVariaveis()
    Dim i As Integer 'data
    Dim j As Integer 'local de cobranca
    Dim k As Integer 'empresa
    
    Dim i2 As Integer
    l_coluna_tab = 0
    l_linha_tab = 1
    l_linha = 0
    
    lQtdData = DateDiff("d", CDate(msk_data_i.Text), CDate(msk_data_f.Text)) + 1
    lQtdLocal = 25
    lQtdEmpresa = 11
    ReDim lTotal(lQtdData, lQtdLocal, lQtdEmpresa)
    For i = 0 To lQtdData
        For j = 0 To lQtdLocal
            For k = 0 To lQtdEmpresa
                lTotal(i, j, k) = 0
            Next
        Next
    Next
    
    'For i = 0 To 11
    '    For i2 = 0 To 10
    '        l_total_1(i, i2) = 0
    '        l_total_2(i, i2) = 0
    '        l_total_3(i, i2) = 0
    '        l_total_4(i, i2) = 0
    '        l_total_5(i, i2) = 0
    '        l_total_6(i, i2) = 0
    '        l_total_7(i, i2) = 0
    '        l_total_8(i, i2) = 0
    '        l_total_9(i, i2) = 0
    '        l_total_10(i, i2) = 0
    '        l_total_11(i, i2) = 0
    '    Next
    'Next
    For i = 1 To 11
        lCodigoEmpresa(i) = 99
        lNomeEmpresa(i) = Space(20)
    Next
    
    If GrupoEmpresa.LocalizarGrupo(g_empresa) Then
        lCodigoGrupoEmpresa = GrupoEmpresa.CodigoGrupo
        lNomeGrupoEmpresa = GrupoEmpresa.NomeGrupo
        If GrupoEmpresa.LocalizarPrimeiro(lCodigoGrupoEmpresa) Then
            i = 1
            lCodigoEmpresa(i) = GrupoEmpresa.CodigoEmpresa
            lNomeEmpresa(i) = GrupoEmpresa.NomeEmpresaReduzido
            Do Until GrupoEmpresa.LocalizarProximo(lCodigoGrupoEmpresa) = False
                i = i + 1
                lCodigoEmpresa(i) = GrupoEmpresa.CodigoEmpresa
                lNomeEmpresa(i) = GrupoEmpresa.NomeEmpresaReduzido
            Loop
        End If
    End If
    
    g_string = ""
End Sub
Private Sub cbo_conta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_fornecedor.SetFocus
    End If
End Sub
Private Sub cbo_fornecedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cboLocalizacaoCobranca.SetFocus
    End If
End Sub
Private Sub cboLocalizacaoCobranca_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_imprimir.SetFocus
    End If
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i.Text = RetiraGString(1)
        msk_data_f.Text = RetiraGString(2)
        cbo_conta.SetFocus
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
    cbo_conta.SetFocus
End Sub
Private Sub cmd_data_i_Click()
    g_string = msk_data_i.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i.Text = RetiraGString(1)
        msk_data_f.Text = RetiraGString(2)
        cbo_conta.SetFocus
    Else
        msk_data_i.Text = RetiraGString(1)
        msk_data_f.SetFocus
    End If
    g_string = ""
End Sub
Private Sub cmd_imprimir_Click()
    If ValidaCampos Then
        If SelecionaImpressoraHP(Me) Then
            Call GravaAuditoria(1, Me.name, 7, "")
            Relatorio
        End If
    End If
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    Dim x_data_i As Date
    Dim x_data_f As Date
    x_data_i = msk_data_i.Text
    x_data_f = msk_data_f.Text
    If Not IsDate(msk_data.Text) Then
        MsgBox "Informe a data de emissão.", vbInformation, "Atenção!"
        msk_data.SetFocus
    ElseIf Not IsDate(msk_data_i.Text) Then
        MsgBox "Informe a data inicial.", vbInformation, "Atenção!"
        msk_data_i.SetFocus
    ElseIf Not IsDate(msk_data_f.Text) Then
        MsgBox "Informe a data final.", vbInformation, "Atenção!"
        msk_data_f.SetFocus
    ElseIf x_data_f < x_data_i Then
        MsgBox "Data final deve ser maior que a data inicial.", vbInformation, "Atenção!"
        msk_data_f.SetFocus
    ElseIf cbo_conta.ListIndex = -1 Then
        MsgBox "Selecione uma conta.", vbInformation, "Atenção!"
        cbo_conta.SetFocus
    ElseIf cbo_fornecedor.ListIndex = -1 Then
        MsgBox "Selecione um fornecedor.", vbInformation, "Atenção!"
        cbo_fornecedor.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_sair_Click()
    Finaliza
    Unload Me
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    If Not IsDate(msk_data.Text) Then
        msk_data.Text = Format(g_data_def, "dd/mm/yyyy")
        msk_data_i.Text = Format(MovimentoContaPagar.PrimeiraDataVencimento(g_empresa), "dd/mm/yyyy")  '  Format(g_data_def, "dd/mm/yyyy")
        msk_data_f.Text = Format(g_data_def + 90, "dd/mm/yyyy")
        cbo_conta.ListIndex = 0
        cbo_fornecedor.ListIndex = 0
        cboLocalizacaoCobranca.ListIndex = 0
        cmd_imprimir.SetFocus
    End If
    Screen.MousePointer = 1
End Sub
Private Sub PreencheCboFornecedor()
    Dim rstFornecedor As New ADODB.Recordset
    
    cbo_fornecedor.Clear
    cbo_fornecedor.AddItem "Todos os Fornecedores"
    cbo_fornecedor.ItemData(cbo_fornecedor.NewIndex) = 0
    Set rstFornecedor = Conectar.RsConexao("SELECT Codigo, Nome FROM Fornecedor WHERE Empresa = " & g_empresa & " ORDER BY Nome")
    'loop RecordSet
    With rstFornecedor
        If Not .BOF Or Not .EOF Then
            .MoveFirst
            Do Until .EOF
                cbo_fornecedor.AddItem !Nome
                cbo_fornecedor.ItemData(cbo_fornecedor.NewIndex) = !Codigo
                .MoveNext
            Loop
        End If
        .Close
    End With
    Set rstFornecedor = Nothing
End Sub
Private Sub PreencheCboLocalizacaoCobranca()
    Dim rstLocalizacao As New ADODB.Recordset
    
    cboLocalizacaoCobranca.Clear
    cboLocalizacaoCobranca.AddItem "Todas as Localizações de Cobrança"
    cboLocalizacaoCobranca.ItemData(cboLocalizacaoCobranca.NewIndex) = 0
    Set rstLocalizacao = Conectar.RsConexao("SELECT Codigo, Nome FROM Local_Cobrancas WHERE Empresa = " & g_empresa & " ORDER BY Nome")
    'loop RecordSet
    With rstLocalizacao
        If Not .BOF Or Not .EOF Then
            .MoveFirst
            Do Until .EOF
                cboLocalizacaoCobranca.AddItem !Nome
                cboLocalizacaoCobranca.ItemData(cboLocalizacaoCobranca.NewIndex) = !Codigo
                .MoveNext
            Loop
        End If
        .Close
    End With
    Set rstLocalizacao = Nothing
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        cmd_imprimir_Click
    End If
End Sub
Private Sub Form_Load()
    CentraForm Me
    PreencheCboConta
    PreencheCboFornecedor
    PreencheCboLocalizacaoCobranca
    l_local = 1
    lRSCriado = False
End Sub
Private Sub msk_data_f_GotFocus()
    msk_data_f.SelStart = 0
    msk_data_f.SelLength = 5
End Sub
Private Sub msk_data_f_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_conta.SetFocus
    End If
End Sub
Private Sub msk_data_GotFocus()
    msk_data.SelStart = 0
    msk_data.SelLength = 5
End Sub
Private Sub msk_data_i_GotFocus()
    msk_data_i.SelStart = 0
    msk_data_i.SelLength = 5
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
