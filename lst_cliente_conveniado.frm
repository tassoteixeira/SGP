VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form lst_cliente_conveniado 
   Caption         =   "Emissão dos Clientes Conveniados"
   ClientHeight    =   2355
   ClientLeft      =   3990
   ClientTop       =   2010
   ClientWidth     =   6795
   Icon            =   "lst_cliente_conveniado.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "lst_cliente_conveniado.frx":030A
   ScaleHeight     =   2355
   ScaleWidth      =   6795
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1140
      Picture         =   "lst_cliente_conveniado.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Visualiza os clientes conveniados em ordem alfabética."
      Top             =   1380
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3000
      Picture         =   "lst_cliente_conveniado.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Imprime os clientes conveniados em ordem alfabética."
      Top             =   1380
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4860
      Picture         =   "lst_cliente_conveniado.frx":3074
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   1380
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      DragMode        =   1  'Automatic
      Height          =   1155
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6555
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   3060
         Picture         =   "lst_cliente_conveniado.frx":4706
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.Data dta_convenio 
         Caption         =   "dta_convenio"
         Connect         =   "Access"
         DatabaseName    =   "Sgp_data.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   3780
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Convenio"
         Top             =   660
         Visible         =   0   'False
         Width           =   2295
      End
      Begin MSMask.MaskEdBox msk_data 
         Height          =   315
         Left            =   1980
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
      Begin MSDBCtls.DBCombo dbcbo_convenio 
         Bindings        =   "lst_cliente_conveniado.frx":59E0
         Height          =   315
         Left            =   1980
         TabIndex        =   5
         Top             =   660
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         Caption         =   "Convenio"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   660
         Width           =   1635
      End
      Begin VB.Label Label5 
         Caption         =   "Data de &emissão"
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1755
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   1620
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "lst_cliente_conveniado"
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
Dim tbl_cliente_conveniado As Table
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    tbl_cliente_conveniado.Close
End Sub
Private Sub ZeraVariaveis()
    lLinha = 0
    lPagina = 0
End Sub
Private Sub Relatorio()
    ZeraVariaveis
    If tbl_cliente_conveniado.RecordCount > 0 Then
        ImpDados
    End If
    cmd_sair.SetFocus
End Sub
Private Sub ImpDados()
    Dim x_linha As String * 80
    'loop cliente conveniado
    With tbl_cliente_conveniado
        tbl_cliente_conveniado.MoveFirst
        Do Until .EOF
            If ![Codigo do Convenio] = Val(dbcbo_convenio.BoundText) Then
                If lPagina = 0 Then
                    ImpCab
                End If
                If lLinha >= 64 Then
                    x_linha = "+---------+------------------------------------------+-------------------------+"
                    Mid(x_linha, 15, 22) = " Cerrado Informática. "
                    BioImprime "@Printer.Print " & x_linha
                    BioImprime "@@Printer.NewPage"
                    ImpCab
                End If
                ImpDet
            End If
            .MoveNext
        Loop
    End With
    ImpTotal
    BioImprime "@@Printer.EndDoc"
    BioFechaImprime
    g_string = lLocal & lNomeArquivo & "|@|Relatório de Cliente Conveniado|@|"
    frm_preview.Show 1
End Sub
Private Sub ImpDet()
    Dim x_linha As String * 80
    Dim i As Integer
    BioImprime "@@Printer.FontName = Draft 10cpi"
    With tbl_cliente_conveniado
        x_linha = "|  CÓDIGO | NOME DO CONVENIADO                       |                         |"
        i = Len(Format(![Codigo do Conveniado], "000,000"))
        Mid(x_linha, 3 + 7 - i, 7) = Format(![Codigo do Conveniado], "000,000")
        Mid(x_linha, 13, 40) = !Nome
        BioImprime "@Printer.Print " & x_linha
    End With
    lLinha = lLinha + 1
End Sub
Private Sub ImpTotal()
    Dim x_linha As String * 80
    Printer.FontName = "Draft 10cpi"
    x_linha = "+---------+------------------------------------------+-------------------------+"
    Mid(x_linha, 15, 22) = " Cerrado Informática. "
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & " "
End Sub
Private Sub ImpCab()
    Dim x_linha As String * 80
    Dim x_string_40 As String * 40
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
    x_linha = "|                                                                  Página,     |"
    x_string_40 = g_nome_empresa
    Mid(x_linha, 3, 40) = x_string_40
    Mid(x_linha, 76, 3) = Format(lPagina, "000")
    BioImprime "@@Printer.FontBold = True"
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontBold = False"
    x_linha = "| RELAÇÃO DE CLIENTES CONVENIADOS EM ORDEM ALFABÉTICA      Goiânia,            |"
    Mid(x_linha, 69, 10) = msk_data
    BioImprime "@Printer.Print " & x_linha
    x_linha = "| CONVÊNIO.:                                                                   |"
    Mid(x_linha, 14, 30) = dbcbo_convenio
    BioImprime "@Printer.Print " & x_linha
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@@Printer.FontBold = False"
    BioImprime "@Printer.Print " & "+---------+------------------------------------------+-------------------------+"
    BioImprime "@Printer.Print " & "|  CÓDIGO | NOME DO CONVENIADO                       |                         |"
    BioImprime "@Printer.Print " & "+---------+------------------------------------------+-------------------------+"
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    msk_data = RetiraGString(1)
    dbcbo_convenio.SetFocus
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
    ElseIf Val(dbcbo_convenio.BoundText) = 0 Then
        MsgBox "Escolha o convênio.", 64, "Atenção!"
        dbcbo_convenio.SetFocus
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
Private Sub dbcbo_convenio_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        cmd_imprimir.SetFocus
    End If
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    If Not IsDate(msk_data.Text) Then
        msk_data.Text = Format(g_data_def, "dd/mm/yyyy")
        dbcbo_convenio.BoundText = 26
        dbcbo_convenio.SetFocus
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
    Set tbl_cliente_conveniado = bd_sgp.OpenTable("Cliente_Conveniado")
    tbl_cliente_conveniado.Index = "id_nome"
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dbcbo_convenio.SetFocus
    End If
End Sub
