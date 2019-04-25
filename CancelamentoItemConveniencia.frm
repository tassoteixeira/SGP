VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form CancelamentoItemConveniencia 
   Caption         =   "Cancelamento de Ítem de Conveniência"
   ClientHeight    =   7590
   ClientLeft      =   1410
   ClientTop       =   1545
   ClientWidth     =   10950
   Icon            =   "CancelamentoItemConveniencia.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   Picture         =   "CancelamentoItemConveniencia.frx":030A
   ScaleHeight     =   7590
   ScaleWidth      =   10950
   Begin VB.Frame frmDados 
      Height          =   7455
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   10815
      Begin VB.CommandButton cmd_ok 
         Caption         =   "&Ok"
         Height          =   855
         Left            =   9060
         Picture         =   "CancelamentoItemConveniencia.frx":0750
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Inicia a pesquisa selecionada."
         Top             =   6480
         Width           =   795
      End
      Begin VB.CommandButton cmd_sair 
         Cancel          =   -1  'True
         Caption         =   "&Sair"
         Height          =   855
         Left            =   9960
         Picture         =   "CancelamentoItemConveniencia.frx":1D5A
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Sai e fecha esta janela."
         Top             =   6480
         Width           =   795
      End
      Begin VB.TextBox txt_celula 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   4800
         TabIndex        =   5
         Top             =   2940
         Visible         =   0   'False
         Width           =   1575
      End
      Begin MSAdodcLib.Adodc oldadodc_venda_conveniencia 
         Height          =   330
         Left            =   5520
         Top             =   3900
         Visible         =   0   'False
         Width           =   3195
         _ExtentX        =   5636
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
         Caption         =   "adodc_venda_conveniencia"
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
      Begin MSFlexGridLib.MSFlexGrid fgd_composicao_caixa 
         Height          =   5895
         Left            =   0
         TabIndex        =   4
         Top             =   540
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   10398
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         BackColorBkg    =   12632256
         AllowUserResizing=   1
      End
      Begin VB.Label lblNumeroCupom 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   6180
         TabIndex        =   9
         Top             =   180
         Width           =   1095
      End
      Begin VB.Label lblData 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   2280
         TabIndex        =   8
         Top             =   180
         Width           =   1095
      End
      Begin VB.Label lbl_total 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   4200
         TabIndex        =   7
         Top             =   6660
         Width           =   1395
      End
      Begin VB.Label Label6 
         Caption         =   "Total do Cupom"
         Height          =   315
         Left            =   2760
         TabIndex        =   6
         Top             =   6660
         Width           =   1395
      End
      Begin VB.Label Label3 
         Caption         =   "Data do Cupom Fiscal"
         Height          =   300
         Index           =   6
         Left            =   120
         TabIndex        =   2
         Top             =   180
         Width           =   2115
      End
      Begin VB.Label Label3 
         Caption         =   "Número do Cupom Fiscal"
         Height          =   300
         Index           =   4
         Left            =   4080
         TabIndex        =   1
         Top             =   180
         Width           =   2055
      End
   End
End
Attribute VB_Name = "CancelamentoItemConveniencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lFlagMovimento As Integer
Dim lOpcao As String
Dim lQtdPeriodo As Integer
Dim lItemMarcado As Integer
Dim lCancelado(0 To 99) As BookmarkEnum

Dim lEmpresa As Integer
Dim lData As Date
Dim lNumero As Long

Const NovaLinha As String = ">*"      ' Indica uma nova linha
Private ControlVisible As Boolean     ' Se o controle esta visivel ou nao
Private LastRow As Long               ' Ultima linha em que se editou
Private LastCol As Long               ' ultima coluna em que se editou
Dim lMarcaCelula As Boolean

Private rsVendaConveniencia As New ADODB.Recordset

Private Sub AtribuiValorCelula()
    Dim Texto As String
    '
    txt_celula.Visible = False
    ControlVisible = False
    '
    ' atribuir o texto anterior a celula
    Select Case LastCol
      Case 4 To 7
        'notas menores que 5 muda cor fonte para vermelho, demais azul
        Texto = txt_celula.Text
        fgd_composicao_caixa.TextMatrix(LastRow, LastCol) = Texto
        'If Val(fgd_composicao_caixa.Text) < 6 Then
        '     fgd_composicao_caixa.CellForeColor = vbRed
        'Else
        '     fgd_composicao_caixa.CellForeColor = vbBlue
        'End If
      Case Else
        'If LastRow = 0 And LastCol = 0 Then
            LastRow = fgd_composicao_caixa.Row
            LastCol = fgd_composicao_caixa.Col
        'End If
      
        Texto = txt_celula.Text
        fgd_composicao_caixa.TextMatrix(LastRow, LastCol) = Texto
    End Select
End Sub
Private Sub AtualizaGrid()
    Dim xTotal As Currency
    Dim i As Integer
    Dim i2 As Integer
    Dim xSQL As String
    
    LimpaGrid
    i = 0
    fgd_composicao_caixa.Visible = False
    xTotal = 0
    
    xSQL = ""
    xSQL = xSQL & "   SELECT Movimento_Venda_Conveniencia.Ordem, Movimento_Venda_Conveniencia.[Codigo do Produto], "
    xSQL = xSQL & "          Movimento_Venda_Conveniencia.[Valor Unitario], Movimento_Venda_Conveniencia.Quantidade, "
    xSQL = xSQL & "          Movimento_Venda_Conveniencia.[Valor Total], Movimento_Venda_Conveniencia.[Item Cancelado], "
    xSQL = xSQL & "          Movimento_Venda_Conveniencia.[Cupom Cancelado], Produto.Nome AS NomeProduto"
    xSQL = xSQL & "     FROM Movimento_Venda_Conveniencia, Produto"
    xSQL = xSQL & "    WHERE Movimento_Venda_Conveniencia.Empresa = " & g_empresa
    xSQL = xSQL & "      AND Movimento_Venda_Conveniencia.Data = " & preparaData(lData)
    xSQL = xSQL & "      AND Movimento_Venda_Conveniencia.[Numero do Cupom] = " & lNumero
    xSQL = xSQL & "      AND Produto.Codigo = Movimento_Venda_Conveniencia.[Codigo do Produto]"
    xSQL = xSQL & " ORDER BY Movimento_Venda_Conveniencia.Ordem"
    Set rsVendaConveniencia = New ADODB.Recordset
    Set rsVendaConveniencia = Conectar.RsConexao(xSQL)
    If Not rsVendaConveniencia.EOF Then
        Do Until rsVendaConveniencia.EOF
            i = i + 1
            fgd_composicao_caixa.Rows = fgd_composicao_caixa.Rows + 1
            fgd_composicao_caixa.Row = i
            fgd_composicao_caixa.Col = 0
            fgd_composicao_caixa.Text = Format(rsVendaConveniencia("Ordem").Value, "###,##0")
            fgd_composicao_caixa.Col = 1
            fgd_composicao_caixa.Text = Format(rsVendaConveniencia("Codigo do Produto").Value, "###,##0")
            fgd_composicao_caixa.Col = 2
            fgd_composicao_caixa.Text = rsVendaConveniencia("NomeProduto").Value
            fgd_composicao_caixa.Col = 3
            fgd_composicao_caixa.Text = Format(rsVendaConveniencia("Valor Unitario").Value, "###,###,##0.00")
            fgd_composicao_caixa.Col = 4
            fgd_composicao_caixa.Text = Format(rsVendaConveniencia("Quantidade").Value, "###,###,##0.00")
            fgd_composicao_caixa.Col = 5
            fgd_composicao_caixa.Text = Format(rsVendaConveniencia("Valor Total").Value, "###,###,##0.00")
            fgd_composicao_caixa.Col = 6
            If rsVendaConveniencia("Item Cancelado").Value Or rsVendaConveniencia("Cupom Cancelado").Value Then
                lCancelado(i) = True
                fgd_composicao_caixa.Text = "Cancelado"
            Else
                lCancelado(i) = False
                fgd_composicao_caixa.Text = "Normal"
                xTotal = xTotal + rsVendaConveniencia("Valor Total").Value
            End If
            rsVendaConveniencia.MoveNext
        Loop
    End If
    
    
    rsVendaConveniencia.Close
    Set rsVendaConveniencia = Nothing
    fgd_composicao_caixa.Row = 1
    fgd_composicao_caixa.Col = 6
    fgd_composicao_caixa.Rows = fgd_composicao_caixa.Rows - 1
    fgd_composicao_caixa.Visible = True
    lbl_total.Caption = Format(xTotal, "###,###,##0.00")
    fgd_composicao_caixa.SetFocus
End Sub
Private Sub ExibirCelula()
    Static OK As Boolean
    '
    ' Se for celula fixa , sair
    If fgd_composicao_caixa.Col <= fgd_composicao_caixa.FixedCols - 1 Or fgd_composicao_caixa.Row <= fgd_composicao_caixa.FixedRows - 1 Then
       Exit Sub
    End If
    
    If OK Then Exit Sub
    OK = True
    '
    txt_celula.Visible = False
    '
    LastRow = fgd_composicao_caixa.Row
    LastCol = fgd_composicao_caixa.Col
    If LastCol = 0 Then
        txt_celula.MaxLength = 4
    ElseIf LastCol = 2 Then
        txt_celula.MaxLength = 10
    End If
    
    '
    ' Nova Celula
    'With fgd_composicao_caixa
    '  If .TextMatrix(LastRow, 0) = NovaLinha Then
    '    .Rows = .Rows + 1
    '    .TextMatrix(LastRow, 0) = LastRow
    '    .TextMatrix(.Rows - 1, 0) = NovaLinha
    '  End If
    'End With
    '
    Select Case LastCol
        Case Else
        txt_celula.Move fgd_composicao_caixa.CellLeft - Screen.TwipsPerPixelX, fgd_composicao_caixa.CellTop + 550 - Screen.TwipsPerPixelY, fgd_composicao_caixa.CellWidth + Screen.TwipsPerPixelX * 2, fgd_composicao_caixa.CellHeight + Screen.TwipsPerPixelY * 2
        txt_celula.Text = fgd_composicao_caixa.Text
        'If Len(fgd_composicao_caixa.Text) = 0 Then
        '   If LastRow > 1 Then
        '       txt_celula.Text = fgd_composicao_caixa.TextMatrix(LastRow - 1, LastCol)
        '   End If
        'End If
        txt_celula.Visible = True
        If txt_celula.Visible Then
          txt_celula.ZOrder
          txt_celula.SetFocus
        End If
    End Select
    ControlVisible = True
    OK = False
End Sub
Private Sub Finaliza()
End Sub
Private Sub LimpaGrid()
    Dim x_sql As String
    Dim i As Integer
    fgd_composicao_caixa.WordWrap = True
    fgd_composicao_caixa.Rows = 2
    fgd_composicao_caixa.Row = 1
    For i = 0 To 6
        fgd_composicao_caixa.Col = i
        fgd_composicao_caixa.Text = ""
    Next
    fgd_composicao_caixa.RowHeight(0) = 500
    fgd_composicao_caixa.Row = 0
    
    i = 0
    fgd_composicao_caixa.Col = i
    fgd_composicao_caixa.Text = "Ordem"
    fgd_composicao_caixa.ColWidth(i) = 600
    fgd_composicao_caixa.ColAlignment(i) = 7
    i = i + 1
    fgd_composicao_caixa.Col = i
    fgd_composicao_caixa.Text = "Código do Produto"
    fgd_composicao_caixa.ColWidth(i) = 900
    fgd_composicao_caixa.ColAlignment(i) = 7
    i = i + 1
    fgd_composicao_caixa.Col = i
    fgd_composicao_caixa.Text = "Nome do Produto"
    fgd_composicao_caixa.ColWidth(i) = 2900
    fgd_composicao_caixa.ColAlignment(i) = 1
    i = i + 1
    fgd_composicao_caixa.Col = i
    fgd_composicao_caixa.Text = "Valor Unitário"
    fgd_composicao_caixa.ColWidth(i) = 900
    fgd_composicao_caixa.ColAlignment(i) = 7
    i = i + 1
    fgd_composicao_caixa.Col = i
    fgd_composicao_caixa.Text = "Quantidade"
    fgd_composicao_caixa.ColWidth(i) = 900
    fgd_composicao_caixa.ColAlignment(i) = 7
    i = i + 1
    fgd_composicao_caixa.Col = i
    fgd_composicao_caixa.Text = "Valor Total"
    fgd_composicao_caixa.ColWidth(i) = 900
    fgd_composicao_caixa.ColAlignment(i) = 7
    i = i + 1
    fgd_composicao_caixa.Col = i
    fgd_composicao_caixa.Text = "Situação"
    fgd_composicao_caixa.ColWidth(i) = 900
    fgd_composicao_caixa.ColAlignment(i) = 4
    'x'lbl_total_nota.Caption = ""
    txt_celula.Visible = False
    fgd_composicao_caixa.Row = 1
    fgd_composicao_caixa.Col = 0
    fgd_composicao_caixa.Text = ""
End Sub
Private Sub cmd_ok_Click()
    If lItemMarcado > 0 Then
        g_string = lItemMarcado & "|@|"
    End If
    cmd_sair_Click
End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub ProximaCelula()
    If fgd_composicao_caixa.Col < fgd_composicao_caixa.Cols - 2 Then
        fgd_composicao_caixa.Col = LastCol + 1
    Else
        fgd_composicao_caixa.Col = 6
        If fgd_composicao_caixa.Row >= fgd_composicao_caixa.Rows - 1 Then
            fgd_composicao_caixa.Row = fgd_composicao_caixa.Row - 1
        End If
        fgd_composicao_caixa.Row = fgd_composicao_caixa.Row + 1
    End If
    'fgd_composicao_caixa.SetFocus
    cmd_ok.SetFocus
End Sub
Private Sub TotalizaGrid()
    Dim x_total As Currency
    Dim i As Integer
    
    x_total = 0
    With fgd_composicao_caixa
        For i = 1 To (.Rows - 1)
            If Len(.TextMatrix(i, 0)) > 0 Then
                If .TextMatrix(i, 6) = "Normal" Then
                    x_total = x_total + fValidaValor(.TextMatrix(i, 5))
                End If
            End If
        Next
    End With
    lbl_total.Caption = Format(x_total, "###,###,##0.00")
End Sub
Private Sub fgd_composicao_caixa_Click()
    ' Quando clicar uma vez
    ' atribui o valor selecionado
    lMarcaCelula = True
    If fgd_composicao_caixa.Col = 6 Then
        LastRow = fgd_composicao_caixa.Row
        LastCol = fgd_composicao_caixa.Col
        txt_celula.Visible = False
    End If
    'AtribuiValorCelula
End Sub
Private Sub fgd_composicao_caixa_DblClick()
    'editar ao clicar duas vezes
    lMarcaCelula = True
    If fgd_composicao_caixa.Col = 6 Then
        '8 - Cancelado/Normal
        LastRow = fgd_composicao_caixa.Row
        LastCol = fgd_composicao_caixa.Col
        txt_celula.Visible = False
        ExibirCelula
    End If
End Sub
Private Sub fgd_composicao_caixa_KeyPress(KeyAscii As Integer)
    lMarcaCelula = True
    Select Case KeyAscii
    ' Editar ao teclar ENTER
    Case vbKeyReturn
        KeyAscii = 0
        If fgd_composicao_caixa.Col = 6 Then
            ExibirCelula
        End If
    ' Cancelar ao pressionar ESC
    Case vbKeyEscape
        KeyAscii = 0
        AtribuiValorCelula
    ' Editar ao pressinar qualquer tecla
    Case 32 To 255
        lMarcaCelula = False
        If fgd_composicao_caixa.Col = 6 Then
            ExibirCelula
            With txt_celula
                If .Visible Then
                    .Text = Chr$(KeyAscii)
                    .SelStart = Len(.Text) + 1
                End If
            End With
        End If
    End Select
End Sub
Private Sub fgd_composicao_caixa_Scroll()
    ' Ver se a coluna esta visivel
    ' entao ocultar os controles
    '
    If fgd_composicao_caixa.ColIsVisible(LastCol) = False Then
        txt_celula.Visible = False
        Exit Sub
    End If
    If fgd_composicao_caixa.RowIsVisible(LastRow) = False Then
        txt_celula.Visible = False
        Exit Sub
    End If
    ' ver se estava visivel antes de ocultar
    ' e posicionar na mesma celula
    If ControlVisible Then
        ExibirCelula
    End If
End Sub
Private Sub Form_Activate()
    If g_empresa <> lEmpresa Then
        lFlagMovimento = 0
    End If
    If lFlagMovimento = 0 Then
        lEmpresa = g_empresa
        lblData.Caption = Format(lData, "dd/mm/yyyy")
        lblNumeroCupom.Caption = Format(lNumero, "###,##0")
        AtualizaGrid
    Else
        lFlagMovimento = 0
    End If
    Screen.MousePointer = 1
End Sub
Private Sub Form_Deactivate()
    lFlagMovimento = 1
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        cmd_sair_Click
    End If
End Sub
Private Sub Form_Load()
    CentraForm Me
    lData = CDate(RetiraGString(1))
    lNumero = CLng(RetiraGString(2))
    g_string = ""
    lItemMarcado = 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub txt_celula_GotFocus()
    With txt_celula
        If LastCol = 6 Then
            .MaxLength = 9
        End If
        If lMarcaCelula Then
            If .Text = "Cancelado" Then
                .Text = "X"
            Else
                .Text = " "
            End If
            .SelStart = 0
            .SelLength = Len(.Text)
        End If
    End With
End Sub
Private Sub txt_celula_KeyPress(KeyAscii As Integer)
    ' ao pressionar ENTER aceitar a entrada de dados
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        fgd_composicao_caixa.SetFocus
    ' ESC, cancela a edição
    ElseIf KeyAscii = vbKeyEscape Then
        KeyAscii = 0
        txt_celula.Visible = False
        ControlVisible = False
    End If
    'If LastCol = 0 Then
    '    Call ValidaInteiro(KeyAscii)
    'ElseIf LastCol = 2 Then
    '    If KeyAscii = 46 Then
    '        KeyAscii = 44
    '    End If
    '    Call ValidaValor(KeyAscii)
    'End If
End Sub
Private Sub txt_celula_LostFocus()
    'Cancelado/Normal
    If LastCol = 6 Then
        If UCase(txt_celula.Text) = "X" Then
            If lItemMarcado = 0 Then
                If lCancelado(fgd_composicao_caixa.Row + 1) = True Then
                    txt_celula.Text = "Normal"
                    MsgBox "Este ítem já está cancelado!", vbInformation, "Cancelamento não aceito!"
                Else
                    lItemMarcado = Val(fgd_composicao_caixa.TextMatrix(fgd_composicao_caixa.Row, 0))
                End If
                txt_celula.Text = "Cancelado"
    '            g_string = "Cancela Ítem de Cupom Fiscal" & "|@|"
    '            g_string = g_string & Me.name & "|@|" & lCodigoFuncionario & "|@|"
                AtribuiValorCelula
            Else
                If lItemMarcado = Val(fgd_composicao_caixa.TextMatrix(fgd_composicao_caixa.Row, 0)) Then
                    MsgBox "Este ítem já está marcado para ser cancelado!", vbInformation, "Ítem já Selecionado!"
                    txt_celula.Text = "Cancelado"
                Else
                    MsgBox "Já tem ítem marcado para cancelamento!" & Chr(10) & "Só poderá ser cancelado 1 ítem por vez.", vbInformation, "Cancelamento não aceito!"
                    If lCancelado(fgd_composicao_caixa.Row + 1) = True Then
                        txt_celula.Text = "Cancelado"
                    Else
                        txt_celula.Text = "Normal"
                    End If
                End If
                AtribuiValorCelula
            End If
        Else
            If lCancelado(fgd_composicao_caixa.Row + 1) = True Then
                txt_celula.Text = "Normal"
                MsgBox "ítem já está cancelado"
            Else
                If lItemMarcado = Val(fgd_composicao_caixa.TextMatrix(fgd_composicao_caixa.Row, 0)) Then
                    lItemMarcado = 0
                End If
            End If
            txt_celula.Text = "Normal"
            AtribuiValorCelula
        End If
        TotalizaGrid
    End If
    If LastCol <> 1 Then
        ProximaCelula
    End If
End Sub
