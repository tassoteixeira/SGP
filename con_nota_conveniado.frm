VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form consulta_nota_conveniado 
   Caption         =   "Consulta Notas de Abastecimento por Conveniado"
   ClientHeight    =   6270
   ClientLeft      =   1455
   ClientTop       =   1785
   ClientWidth     =   7830
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "con_nota_conveniado.frx":0000
   ScaleHeight     =   6270
   ScaleWidth      =   7830
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   6960
      Picture         =   "con_nota_conveniado.frx":0046
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   120
      Width           =   795
   End
   Begin VB.Frame frmDados 
      Height          =   2055
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6855
      Begin MSAdodcLib.Adodc adodcClienteConveniado 
         Height          =   330
         Left            =   3480
         Top             =   660
         Visible         =   0   'False
         Width           =   2580
         _ExtentX        =   4551
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
         Caption         =   "adodcClienteConveniado"
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
      Begin MSAdodcLib.Adodc adodcCliente 
         Height          =   330
         Left            =   3420
         Top             =   180
         Visible         =   0   'False
         Width           =   2580
         _ExtentX        =   4551
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
      Begin VB.TextBox txt_conveniado 
         Height          =   285
         Left            =   1260
         MaxLength       =   6
         TabIndex        =   5
         Top             =   600
         Width           =   795
      End
      Begin VB.CheckBox chk_fixar_data_final 
         Caption         =   "&Mantém última data fixa"
         Height          =   255
         Left            =   180
         TabIndex        =   11
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox txt_cliente 
         Height          =   285
         Left            =   1260
         MaxLength       =   6
         TabIndex        =   2
         Top             =   240
         Width           =   795
      End
      Begin MSMask.MaskEdBox msk_data_final 
         Height          =   300
         Left            =   5640
         TabIndex        =   10
         Top             =   1200
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
         Left            =   1260
         TabIndex        =   8
         Top             =   1200
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSDataListLib.DataCombo dtcboCliente 
         Bindings        =   "con_nota_conveniado.frx":16D8
         Height          =   315
         Left            =   2160
         TabIndex        =   3
         Top             =   240
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Razao Social"
         BoundColumn     =   "Codigo"
         Text            =   "dtcboCliente"
      End
      Begin MSDataListLib.DataCombo dtcboClienteConveniado 
         Bindings        =   "con_nota_conveniado.frx":16F3
         Height          =   315
         Left            =   2160
         TabIndex        =   6
         Top             =   600
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Nome"
         BoundColumn     =   "Codigo do Conveniado"
         Text            =   "dtcboClienteConveniado"
      End
      Begin VB.Label Label3 
         Caption         =   "C&onveniado"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   5640
         TabIndex        =   13
         Top             =   1620
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Total"
         Height          =   195
         Left            =   5100
         TabIndex        =   12
         Top             =   1620
         Width           =   435
      End
      Begin VB.Label Label1 
         Caption         =   "&Data Inicial"
         Height          =   255
         Index           =   1
         Left            =   1260
         TabIndex        =   7
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "D&ata Final"
         Height          =   255
         Index           =   2
         Left            =   5640
         TabIndex        =   9
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "C&liente"
         Height          =   315
         Index           =   9
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid 
      Height          =   4095
      Left            =   0
      TabIndex        =   15
      Top             =   2160
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   7223
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
   End
End
Attribute VB_Name = "consulta_nota_conveniado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lDataFinal As Date
Dim lSQL As String
Dim lTotal As Currency
Dim lCliente As Long
Dim lConveniado As Long

Private Cliente As New cCliente
Private ClienteConveniado As New cClienteConveniado
Private MovimentoNotaAbastecimento As New cMovimentoNotaAbastecimento
Private rsTabela As New adodb.Recordset

Private Sub AtualizaMSFlexGrid()
    Dim i As Integer
    
    On Error GoTo ErroConsulta
    
    'Verifica movimento
    i = 0
    If Len(txt_conveniado.Text) = 0 Then
        Exit Sub
    End If
    BuscaDatas
    If rsTabela.RecordCount > 0 Then
        rsTabela.MoveFirst
        Do Until rsTabela.EOF
            If rsTabela![Data do Abastecimento] >= CDate(msk_data_inicial.Text) And rsTabela![Data do Abastecimento] <= CDate(msk_data_final.Text) Then
                MSFlexGrid.Rows = MSFlexGrid.Rows + 1
                i = i + 1
                MSFlexGrid.Row = i
                MSFlexGrid.Col = 0
                MSFlexGrid.Text = rsTabela("Data do Abastecimento").Value
                MSFlexGrid.Col = 1
                MSFlexGrid.Text = rsTabela("Periodo").Value
                MSFlexGrid.Col = 2
                MSFlexGrid.Text = rsTabela("Tipo do Movimento").Value
                MSFlexGrid.Col = 3
                MSFlexGrid.Text = rsTabela("Nome").Value
                MSFlexGrid.Col = 4
                MSFlexGrid.Text = Format(rsTabela("Valor Total").Value, "###,###,##0.00")
                MSFlexGrid.Col = 5
                MSFlexGrid.Text = rsTabela("Numero da Nota").Value
                MSFlexGrid.Col = 6
                MSFlexGrid.Text = rsTabela("Codigo do Produto2").Value
                MSFlexGrid.Col = 7
                MSFlexGrid.Text = rsTabela("Empresa").Value
            End If
            rsTabela.MoveNext
        Loop
    End If
    rsTabela.Close
    Exit Sub
    
ErroConsulta:
    If Err = 3075 Then
        MsgBox "Condição inválida.", vbExclamation, "Erro de Consulta"
    Else
        MsgBox Error, vbExclamation, "Erro de Consulta"
    End If
    Exit Sub
End Sub
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    Set Cliente = Nothing
    Set ClienteConveniado = Nothing
    Set MovimentoNotaAbastecimento = Nothing
End Sub
Private Sub LimpaMSFlexGrid()
    Dim i As Integer
    
    MSFlexGrid.WordWrap = True
    MSFlexGrid.Rows = 2
    MSFlexGrid.Row = 1
    For i = 0 To 7
        MSFlexGrid.Col = i
        MSFlexGrid.Text = ""
    Next
    MSFlexGrid.RowHeight(0) = 500
    MSFlexGrid.Row = 0
    i = 0
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Data abast."
    MSFlexGrid.ColWidth(i) = 1000
    MSFlexGrid.ColAlignment(i) = 4
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Per."
    MSFlexGrid.ColWidth(i) = 400
    MSFlexGrid.ColAlignment(i) = 4
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Tipo mov."
    MSFlexGrid.ColWidth(i) = 400
    MSFlexGrid.ColAlignment(i) = 4
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Nome do Produto"
    MSFlexGrid.ColWidth(i) = 3000
    MSFlexGrid.ColAlignment(i) = 1
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Valor Total"
    MSFlexGrid.ColWidth(i) = 800
    MSFlexGrid.ColAlignment(i) = 4
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Numero da Nota"
    MSFlexGrid.ColWidth(i) = 800
    MSFlexGrid.ColAlignment(i) = 4
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Codigo do Produto"
    MSFlexGrid.ColWidth(i) = 900
    MSFlexGrid.ColAlignment(i) = 4
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Empresa"
    MSFlexGrid.ColWidth(i) = 750
    MSFlexGrid.ColAlignment(i) = 4
    MSFlexGrid.Row = 1
    MSFlexGrid.Col = 0
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If dtcboCliente.BoundText = "" Then
        MsgBox "Escolha o cliente.", vbInformation, "Atenção!"
        dtcboCliente.SetFocus
    ElseIf Not IsDate(msk_data_inicial.Text) Then
        MsgBox "Informe a data inicial.", vbInformation, "Atenção!"
        msk_data_inicial.SetFocus
    ElseIf Not IsDate(msk_data_final.Text) Then
        MsgBox "Informe a data final.", vbInformation, "Atenção!"
        msk_data_final.SetFocus
    ElseIf CDate(msk_data_final.Text) < CDate(msk_data_inicial.Text) Then
        MsgBox "A data final dever ser maior que " & CDate(msk_data_inicial.Text) - 1 & ".", vbInformation, "Atenção!"
        msk_data_final.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub chk_fixar_data_final_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dtcboCliente_LostFocus
    End If
End Sub
Private Sub cmd_sair_Click()
    Finaliza
    Unload Me
End Sub
Function BuscaDatas() As Boolean
    BuscaDatas = False
    lTotal = 0
    msk_data_inicial.Text = "__/__/____"
    msk_data_final.Text = "__/__/____"
    lblTotal.Caption = ""
    
    lSQL = "SELECT Movimento_Nota_Abastecimento.[Data do Abastecimento], Movimento_Nota_Abastecimento.Periodo, Movimento_Nota_Abastecimento.[Tipo do Movimento], Produto.Nome, Movimento_Nota_Abastecimento.[Valor Total], Movimento_Nota_Abastecimento.[Numero da Nota], Movimento_Nota_Abastecimento.[Codigo do Produto2], Movimento_Nota_Abastecimento.Empresa"
    lSQL = lSQL & "  FROM Movimento_Nota_Abastecimento, Produto"
    lSQL = lSQL & " WHERE Produto.Codigo = Movimento_Nota_Abastecimento.[Codigo do Produto2]"
    lSQL = lSQL & "   AND Movimento_Nota_Abastecimento.[Codigo do Cliente] = " & Val(txt_cliente.Text)
    lSQL = lSQL & "   AND Movimento_Nota_Abastecimento.[Codigo do Conveniado] = " & Val(txt_conveniado.Text)
    lSQL = lSQL & " ORDER BY [Data do Abastecimento], Periodo, [Numero da Nota], [Codigo do Produto2]"
    'Abre RecordSet
    Set rsTabela = Conectar.RsConexao(lSQL)
    With rsTabela
        If .RecordCount > 0 Then
            .MoveFirst
            Do Until .EOF
                If Not IsDate(msk_data_inicial.Text) Then
                    msk_data_inicial.Text = Format(![Data do Abastecimento], "dd/mm/yyyy")
                End If
                If Not IsDate(lDataFinal) Then
                    lDataFinal = Format(![Data do Abastecimento], "dd/mm/yyyy")
                End If
                If chk_fixar_data_final Then
                    msk_data_final.Text = Format(lDataFinal, "dd/mm/yyyy")
                Else
                    msk_data_final.Text = Format(![Data do Abastecimento], "dd/mm/yyyy")
                End If
                If ![Data do Abastecimento] >= CDate(msk_data_inicial.Text) And ![Data do Abastecimento] <= CDate(msk_data_final.Text) Then
                    lTotal = lTotal + ![Valor Total]
                End If
                .MoveNext
            Loop
        End If
        If lTotal > 0 Then
            BuscaDatas = True
            lblTotal.Caption = Format(lTotal, "###,###,##0.00")
        Else
            msk_data_inicial.Text = g_data_def
            msk_data_final.Text = g_data_def
            MsgBox "Este cliente conveniado não tem notas de abastecimento!", 48, "Baixa de Notas de Abastecimento."
            'cmd_sair.SetFocus
        End If
    End With
    Call GravaAuditoria(1, Me.name, 5, "Cli:" & CLng(txt_cliente.Text) & " Ref:" & msk_data_inicial.Text & " a " & msk_data_final.Text & " Vlr:" & lblTotal.Caption)
End Function
Private Sub dtcboCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        txt_conveniado.SetFocus
    End If
End Sub
Private Sub dtcboCliente_LostFocus()
    If dtcboCliente.BoundText <> "" Then
        If Cliente.LocalizarCodigo(CLng(dtcboCliente.BoundText)) Then
            txt_cliente.Text = Cliente.Codigo
            Set adodcClienteConveniado.Recordset = Conectar.RsConexao("SELECT [Codigo do Conveniado], Nome FROM Cliente_Conveniado WHERE [Codigo do Convenio] = " & CLng(dtcboCliente.BoundText) & " ORDER BY Nome")
        End If
    End If
End Sub
Private Sub dtcboClienteConveniado_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        msk_data_inicial.SetFocus
    End If
End Sub
Private Sub dtcboClienteConveniado_LostFocus()
    lConveniado = 0
    If dtcboClienteConveniado.BoundText <> "" Then
        If ClienteConveniado.LocalizarCodigo(CLng(txt_cliente.Text), CLng(dtcboClienteConveniado.BoundText)) Then
            txt_conveniado.Text = ClienteConveniado.CodigoConveniado
            lConveniado = CLng(txt_conveniado.Text)
            If BuscaDatas Then
                msk_data_final_LostFocus
            End If
        End If
    End If
End Sub
Private Sub Form_Activate()
    lConveniado = 0
    lCliente = 26
    txt_cliente = lCliente
    dtcboCliente.BoundText = lCliente
    LimpaMSFlexGrid
    txt_conveniado.SetFocus
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
    Set adodcCliente.Recordset = Conectar.RsConexao("SELECT Codigo, [Razao Social] FROM Cliente WHERE Inativo = " & preparaBooleano(False) & " AND [Codigo do Convenio] > 1 ORDER BY [Razao Social]")
End Sub
Private Sub MSFlexGrid_DblClick()
    txt_conveniado.SetFocus
End Sub
Private Sub MSFlexGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        txt_conveniado.SetFocus
    ElseIf KeyCode = 98 Then
        KeyCode = 40
    ElseIf KeyCode = 99 Then
        KeyCode = 34
    ElseIf KeyCode = 104 Then
        KeyCode = 38
    ElseIf KeyCode = 105 Then
        KeyCode = 33
    End If
End Sub
Private Sub MSFlexGrid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
        KeyAscii = 0
        txt_conveniado.SetFocus
    End If
End Sub
Private Sub msk_data_final_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        chk_fixar_data_final.SetFocus
    End If
End Sub
Private Sub msk_data_final_LostFocus()
    If ValidaCampos Then
        lDataFinal = msk_data_final.Text
        AtualizaMSFlexGrid
        BuscaDatas
        MSFlexGrid.SetFocus
    End If
End Sub
Private Sub msk_data_inicial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data_final.SetFocus
    End If
End Sub
Private Sub txt_cliente_GotFocus()
    lTotal = 0
    txt_conveniado.Text = ""
    lblTotal.Caption = ""
    msk_data_inicial.Text = "01/01/1900"
    msk_data_final.Text = "01/01/1900"
    
    AtualizaMSFlexGrid
    msk_data_inicial.Text = "__/__/____"
    msk_data_final.Text = "__/__/____"
End Sub
Private Sub txt_cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dtcboCliente.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_cliente_LostFocus()
    lCliente = 0
    If Val(txt_cliente.Text) > 0 Then
        lCliente = CLng(txt_cliente.Text)
        If Cliente.LocalizarCodigo(CLng(txt_cliente.Text)) Then
            dtcboCliente.BoundText = CLng(txt_cliente.Text)
            dtcboCliente_LostFocus
            Exit Sub
        Else
            MsgBox "Cliente não cadastro.", vbInformation, "Atenção!"
            dtcboCliente.BoundText = ""
            txt_cliente.SetFocus
        End If
    End If
End Sub
Private Sub txt_conveniado_GotFocus()
    lTotal = 0
    txt_conveniado.Text = ""
    lblTotal.Caption = ""
    msk_data_inicial.Text = "01/01/1900"
    msk_data_final.Text = "01/01/1900"
    AtualizaMSFlexGrid
    msk_data_inicial.Text = "__/__/____"
    msk_data_final.Text = "__/__/____"
End Sub
Private Sub txt_conveniado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dtcboClienteConveniado.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_conveniado_LostFocus()
    lConveniado = 0
    If Val(txt_conveniado.Text) > 0 Then
        lConveniado = CLng(txt_conveniado.Text)
        If ClienteConveniado.LocalizarCodigo(CLng(txt_cliente.Text), CLng(txt_conveniado.Text)) Then
            dtcboClienteConveniado.BoundText = CLng(txt_conveniado.Text)
            dtcboClienteConveniado_LostFocus
            Exit Sub
        Else
            MsgBox "Cliente Conveniado não cadastro.", vbInformation, "Atenção!"
            dtcboClienteConveniado.BoundText = ""
            txt_conveniado.SetFocus
        End If
    End If
End Sub

