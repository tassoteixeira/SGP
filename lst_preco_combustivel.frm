VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form emissao_preco_combustivel 
   Caption         =   "Emissão dos Preços de Combustíveis"
   ClientHeight    =   1875
   ClientLeft      =   1350
   ClientTop       =   1680
   ClientWidth     =   4905
   Icon            =   "lst_preco_combustivel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1875
   ScaleWidth      =   4905
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   660
      Picture         =   "lst_preco_combustivel.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Visualiza preços de combustíveis."
      Top             =   900
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   2040
      Picture         =   "lst_preco_combustivel.frx":1A24
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Imprime preços de combustíveis."
      Top             =   900
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3420
      Picture         =   "lst_preco_combustivel.frx":302E
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   900
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   4635
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2820
         Picture         =   "lst_preco_combustivel.frx":46C0
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
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label5 
         Caption         =   "Data de &emissão"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   1515
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "emissao_preco_combustivel"
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
Dim l_preco_a As Currency
Dim l_preco_aa As Currency
Dim l_preco_d As Currency
Dim l_preco_da As Currency
Dim l_preco_g As Currency
Dim l_preco_ga As Currency
Dim tbl_bomba As Table
Dim tbl_combustivel As Table
Dim tbl_empresa As Table
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    tbl_bomba.Close
    tbl_combustivel.Close
    tbl_empresa.Close
End Sub
Private Sub ZeraVariaveis()
    Dim i As Integer
    lLinha = 0
    lPagina = 0
End Sub
Private Sub Relatorio()
    ZeraVariaveis
    If tbl_combustivel.RecordCount > 0 Then
        LoopTabelaEmpresa
        ImpRodape
        BioImprime "@@Printer.EndDoc"
        BioFechaImprime
        g_string = lLocal & lNomeArquivo & "|@|Relatório do Preço de Combustível|@|"
        frm_preview.Show 1
    End If
    cmd_sair.SetFocus
End Sub
Private Sub LoopCombustivel(i As Integer)
    l_preco_a = 0
    l_preco_aa = 0
    l_preco_d = 0
    l_preco_da = 0
    l_preco_g = 0
    l_preco_ga = 0
    With tbl_combustivel
        .Seek ">=", i, " ", " "
        If Not .NoMatch Then
            Do Until .EOF
                If !Empresa <> i Then
                    Exit Do
                End If
                tbl_bomba.Seek "=", i, !Codigo
                If Not tbl_bomba.NoMatch Then
                    If Trim(!Codigo) = "A" Then
                        l_preco_a = tbl_bomba![Preco de Custo]
                    ElseIf Trim(!Codigo) = "AA" Then
                        l_preco_aa = tbl_bomba![Preco de Custo]
                    ElseIf Trim(!Codigo) = "D" Then
                        l_preco_d = tbl_bomba![Preco de Custo]
                    ElseIf Trim(!Codigo) = "DA" Then
                        l_preco_da = tbl_bomba![Preco de Custo]
                    ElseIf Trim(!Codigo) = "G" Then
                        l_preco_g = tbl_bomba![Preco de Custo]
                    ElseIf Trim(!Codigo) = "GA" Then
                        l_preco_ga = tbl_bomba![Preco de Custo]
                    End If
                End If
                .MoveNext
            Loop
        End If
    End With
End Sub
Private Sub LoopTabelaEmpresa()
    Dim x_linha As String * 80
    'loop tabela empresa
    With tbl_empresa
        .MoveFirst
        Do Until .EOF
            If !Codigo > 9 Then
                Exit Do
            End If
            If !Codigo <> 1 And !Codigo <> 5 And !Codigo <> 7 And !Codigo <> 8 Then
                Call LoopCombustivel(!Codigo)
                If lPagina = 0 Then
                    ImpCab
                End If
                If lLinha >= 55 Then
                    x_linha = "+------------+----------+----------+----------+----------+----------+----------+"
                    Mid(x_linha, 4, 22) = " Cerrado Informática. "
                    BioImprime "@Printer.Print " & x_linha
                    BioImprime "@@Printer.NewPage"
                    ImpCab
                End If
                Call ImpDet(!Codigo)
            End If
            .MoveNext
        Loop
    End With
End Sub
Private Sub ImpDet(x_empresa As Integer)
    Dim x_linha As String * 137
    Dim i As Integer
    BioImprime "@Printer.Print " & "+------------+----------+----------+----------+----------+----------+----------+"
    BioImprime "@Printer.Print " & "|            |          |          |          |          |          |          |"
    x_linha = "|            |          |          |          |          |          |          |"
    If x_empresa = 1 Then
        Mid(x_linha, 3, 10) = "POSTO 1   "
    ElseIf x_empresa = 2 Then
        Mid(x_linha, 3, 10) = "POSTO 2   "
    ElseIf x_empresa = 3 Then
        Mid(x_linha, 3, 10) = "POSTO 3   "
    ElseIf x_empresa = 4 Then
        Mid(x_linha, 3, 10) = "POSTO 4   "
    ElseIf x_empresa = 5 Then
        Mid(x_linha, 3, 10) = "POSTO 5   "
    ElseIf x_empresa = 6 Then
        Mid(x_linha, 3, 10) = "POSTO 6   "
    ElseIf x_empresa = 9 Then
        Mid(x_linha, 3, 10) = "POSTO 7   "
    End If
    If l_preco_a > 0 Then
        i = Len(Format(l_preco_a, "#,##0.0000"))
        Mid(x_linha, 15 + 10 - i, i) = Format(l_preco_a, "#,##0.0000")
    End If
    If l_preco_aa > 0 Then
        i = Len(Format(l_preco_aa, "#,##0.0000"))
        Mid(x_linha, 26 + 10 - i, i) = Format(l_preco_aa, "#,##0.0000")
    End If
    If l_preco_d > 0 Then
        i = Len(Format(l_preco_d, "#,##0.0000"))
        Mid(x_linha, 37 + 10 - i, i) = Format(l_preco_d, "#,##0.0000")
    End If
    If l_preco_da > 0 Then
        i = Len(Format(l_preco_da, "#,##0.0000"))
        Mid(x_linha, 48 + 10 - i, i) = Format(l_preco_da, "#,##0.0000")
    End If
    If l_preco_g > 0 Then
        i = Len(Format(l_preco_g, "#,##0.0000"))
        Mid(x_linha, 59 + 10 - i, i) = Format(l_preco_g, "#,##0.0000")
    End If
    If l_preco_ga > 0 Then
        i = Len(Format(l_preco_ga, "#,##0.0000"))
        Mid(x_linha, 70 + 10 - i, i) = Format(l_preco_ga, "#,##0.0000")
    End If
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@Printer.Print " & "|            |          |          |          |          |          |          |"
    lLinha = lLinha + 1
End Sub
Private Sub ImpRodape()
    Dim x_linha As String * 80
    x_linha = "+------------+----------+----------+----------+----------+----------+----------+"
    Mid(x_linha, 4, 22) = " Cerrado Informática. "
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & " "
End Sub
Private Sub ImpCab()
    Dim x_linha As String * 137
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
    x_linha = "| GRUPO X                                                          Página, " & Format(lPagina, "000") & " |"
    g_string = ReadINI("GRUPO DE EMPRESAS", "Nome do Grupo", gArquivoIni)
    Mid(x_linha, 3, 40) = g_string
    g_string = ""
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontBold = False"
    BioImprime "@Printer.Print " & "| PREÇO DE CUSTO DOS COMBUSTÍVEIS                          Goiânia, " & msk_data & " |"
    BioImprime "@Printer.Print " & "+------------+----------+----------+----------+----------+----------+----------+"
    BioImprime "@Printer.Print " & "| EMPRESA    | ÁLCOOL   | ÁLCOOL + | DIESEL   | DIESEL + | GASOLINA | GASOLINA+|"
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    msk_data = RetiraGString(1)
    cmd_visualizar.SetFocus
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
        msk_data.Text = Format(Date, "dd/mm/yyyy")
    End If
    msk_data.SetFocus
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
    Set tbl_bomba = bd_sgp.OpenTable("Bomba")
    Set tbl_combustivel = bd_sgp.OpenTable("Combustivel")
    Set tbl_empresa = bd_sgp.OpenTable("Empresas")
    tbl_bomba.Index = "id_combustivel"
    tbl_combustivel.Index = "id_nome"
    tbl_empresa.Index = "id_codigo"
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_visualizar.SetFocus
    End If
End Sub
