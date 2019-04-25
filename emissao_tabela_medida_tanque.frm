VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form emissao_tabela_medida_tanque 
   Caption         =   "Emissão da Tabela de Medida de Tanques"
   ClientHeight    =   3570
   ClientLeft      =   3990
   ClientTop       =   2010
   ClientWidth     =   6795
   Icon            =   "emissao_tabela_medida_tanque.frx":0000
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   Picture         =   "emissao_tabela_medida_tanque.frx":030A
   ScaleHeight     =   3570
   ScaleWidth      =   6795
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1140
      Picture         =   "emissao_tabela_medida_tanque.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Visualiza a tabela de medida de tanque."
      Top             =   2640
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3000
      Picture         =   "emissao_tabela_medida_tanque.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Imprime a tabela de medida de tanque."
      Top             =   2640
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4860
      Picture         =   "emissao_tabela_medida_tanque.frx":3074
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   2640
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6555
      Begin VB.OptionButton optTanque30 
         Caption         =   "Tanque de 30.000 Lts"
         Height          =   315
         Left            =   180
         TabIndex        =   7
         Top             =   1980
         Width           =   2115
      End
      Begin VB.OptionButton optTanque20 
         Caption         =   "Tanque de 20.000 Lts"
         Height          =   315
         Left            =   180
         TabIndex        =   6
         Top             =   1560
         Width           =   2115
      End
      Begin VB.OptionButton optTanque15 
         Caption         =   "Tanque de 15.000 Lts"
         Height          =   315
         Left            =   180
         TabIndex        =   5
         Top             =   1140
         Value           =   -1  'True
         Width           =   2115
      End
      Begin VB.OptionButton optTanque10 
         Caption         =   "Tanque de 10.000 Lts"
         Height          =   315
         Left            =   180
         TabIndex        =   4
         Top             =   720
         Width           =   2115
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2760
         Picture         =   "emissao_tabela_medida_tanque.frx":4706
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
      Left            =   0
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "emissao_tabela_medida_tanque"
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
Dim lMedida(0 To 4) As Integer
Dim lLitro(0 To 4) As Currency
Dim lSQL As String
Dim lRSCriado As Boolean
Private rs As New adodb.Recordset
Dim rs2 As New adodb.Recordset
Private Sub CriaRS()
    With rs2
        If lRSCriado Then
            .MoveFirst
            Do Until .EOF
                .Delete
                .MoveNext
            Loop
        Else
            .CursorLocation = adUseClient
            .Fields.Append "Ordem", adVarChar, 8
            .Fields.Append "Medida", adVarChar, 4
            .Fields.Append "Litros", adVarChar, 10
            .Open
            lRSCriado = True
        End If
    End With
End Sub
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set rs = Nothing
End Sub
Private Sub ZeraVariaveis()
    Dim i As Integer
    lLinha = 0
    lPagina = 0
    For i = 0 To 4
        lMedida(i) = 0
        lLitro(i) = 0
    Next
End Sub
Private Sub Relatorio()
    ZeraVariaveis
    
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "SELECT Medida, "
    If optTanque10.Value Then
        lSQL = lSQL & "[Medicao Tanque 10]"
    ElseIf optTanque15.Value Then
        lSQL = lSQL & "[Medicao Tanque 15]"
    ElseIf optTanque20.Value Then
        lSQL = lSQL & "[Medicao Tanque 20]"
    ElseIf optTanque30.Value Then
        lSQL = lSQL & "[Medicao Tanque 30]"
    End If
    lSQL = lSQL & " AS Quantidade"
    
    lSQL = lSQL & "  FROM Conversao_Medicao_Combustivel"
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    lSQL = lSQL & " ORDER BY Medida"
    
    'Abre RecordSet
    Set rs = New adodb.Recordset
    Set rs = Conectar.RsConexao(lSQL)
    
    
    'Verifica movimento
    If rs.RecordCount > 0 Then
        CriaRS
        GravaRS
        ImpDados
    End If
    If rs.State = 1 Then
        rs.Close
    End If
    cmd_sair.SetFocus
End Sub
Private Sub ImpDados()
    Dim i As Integer
    'loop movimento de cheques
    i = 0
    rs2.Sort = "Ordem"
    rs2.MoveFirst
    Do Until rs2.EOF
        i = i + 1
        lMedida(i) = rs2("Medida").Value
        lLitro(i) = rs2("Litros").Value
        If i = 4 Then
            ImpDet
            For i = 0 To 4
                lMedida(i) = 0
                lLitro(i) = 0
            Next
            i = 0
        End If
        rs2.MoveNext
    Loop
    'ImpSubTotal
    If lPagina > 0 Then
        ImpTotal
        BioImprime "@@Printer.EndDoc"
        BioFechaImprime
        g_string = lLocal & lNomeArquivo & "|@|Relatório de Medição de Combustíveis|@|"
        frm_preview.Show 1
    End If
End Sub
Private Sub ImpDet()
    Dim xLinha As String
    Dim i As Integer
    Dim i2 As Integer
    If lPagina = 0 Then
        ImpCab
    End If
    If lLinha >= 57 Then
        xLinha = "+-------------------+-------------------+-------------------+------------------+"
        Mid(xLinha, 5, 22) = " Cerrado Informática. "
        BioImprime "@Printer.Print " & xLinha
        BioImprime "@@Printer.NewPage"
        ImpCab
    End If
    '                  1         2         3         4         5         6         7         8
    '         12345678901234567890123456789012345678901234567890123456789012345678901234567890
    xLinha = "|       -           |       -           |       -           |       -          |"
    For i2 = 1 To 4
        If lLitro(i2) > 0 Then
            i = Len(Format(lMedida(i2), "##0"))
            Mid(xLinha, (i2 * 20 - 20 + 4) + 3 - i, i) = Format(lMedida(i2), "##0")
            i = Len(Format(lLitro(i2), "##,##0.0"))
            Mid(xLinha, (i2 * 20 - 20 + 12) + 8 - i, i) = Format(lLitro(i2), "##,##0.0")
        End If
    Next
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 1
End Sub
Private Sub ImpTotal()
    Dim xLinha As String
    xLinha = "+-------------------+-------------------+-------------------+------------------+"
    Mid(xLinha, 5, 22) = " Cerrado Informática. "
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & "  "
End Sub
Private Sub GravaRS()
    Dim i As Integer
    Dim i2 As Integer
    
    i = 0
    i2 = rs.RecordCount / 4
    Do Until rs.EOF
        i = i + 1
        If i > i2 Then
            i = 1
        End If
        rs2.AddNew
        rs2("Medida").Value = rs("Medida").Value
        rs2("Ordem").Value = Format(i, "000") & Format(rs("Medida").Value, "000")
        rs2("Litros").Value = Format(rs("Quantidade").Value, "0000000.00")
        rs2.Update
        rs.MoveNext
    Loop
End Sub
Private Sub ImpCab()
    Dim i As Integer
    Dim xLinha As String
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
    xLinha = "|                                                           CIDADE, __/__/____ |"
    Mid(xLinha, 3, 40) = g_nome_empresa
    i = Len(g_cidade_empresa)
    Mid(xLinha, 37 + 30 - i, i) = g_cidade_empresa
    Mid(xLinha, 69, 10) = msk_data.Text
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| TABELA DE CONVERSAO DE MEDIDAS DE TANQUE ( PARA TANQUE DE __.000 LITROS )    |"
    If optTanque10.Value Then
        Mid(xLinha, 61, 2) = "10"
    ElseIf optTanque15.Value Then
        Mid(xLinha, 61, 2) = "15"
    ElseIf optTanque20.Value Then
        Mid(xLinha, 61, 2) = "20"
    ElseIf optTanque30.Value Then
        Mid(xLinha, 61, 2) = "30"
    End If
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontBold = False"
    BioImprime "@Printer.Print " & "+-------------------+-------------------+-------------------+------------------+"
    BioImprime "@Printer.Print " & "|MEDIDA -  LITROS   |MEDIDA -  LITROS   |MEDIDA -  LITROS   |MEDIDA -  LITROS  |"
    BioImprime "@Printer.Print " & "+-------------------+-------------------+-------------------+------------------+"
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    msk_data.Text = RetiraGString(1)
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
    If Not IsDate(msk_data.Text) Then
        MsgBox "Informe a data de emissão.", vbInformation, "Atenção!"
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
        msk_data.Text = Format(g_data_def, "dd/mm/yyyy")
        msk_data.SetFocus
    End If
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
    lRSCriado = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_GotFocus()
    msk_data.SelStart = 0
    msk_data.SelLength = 2
End Sub
Private Sub msk_data_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_visualizar.SetFocus
    End If
End Sub
