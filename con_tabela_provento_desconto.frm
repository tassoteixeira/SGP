VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form consulta_tabela_provento_desconto 
   Caption         =   "Consulta Tabela de Proventos/Descontos"
   ClientHeight    =   5355
   ClientLeft      =   1245
   ClientTop       =   2385
   ClientWidth     =   8655
   Icon            =   "con_tabela_provento_desconto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "con_tabela_provento_desconto.frx":0442
   ScaleHeight     =   5355
   ScaleWidth      =   8655
   Begin VB.Frame Frame1 
      Caption         =   "S&elecionar por..."
      Height          =   1155
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6855
      Begin VB.ComboBox cbo_operador 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txt_condicao 
         Height          =   285
         Left            =   3120
         TabIndex        =   6
         Top             =   720
         Width           =   3615
      End
      Begin VB.ComboBox cbo_campo 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   300
         Width           =   5775
      End
      Begin VB.Label Label4 
         Caption         =   "O&perador"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   780
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Co&ndição"
         Height          =   255
         Left            =   2340
         TabIndex        =   5
         Top             =   780
         Width           =   795
      End
      Begin VB.Label Label2 
         Caption         =   "&Campo"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   675
      End
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   7020
      Picture         =   "con_tabela_provento_desconto.frx":0488
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Inicia a pesquisa selecionada."
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   7860
      Picture         =   "con_tabela_provento_desconto.frx":1A92
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   120
      Width           =   735
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid 
      Height          =   4035
      Left            =   60
      TabIndex        =   9
      Top             =   1260
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   7117
      _Version        =   393216
      FixedCols       =   0
      BackColorSel    =   12582912
      ForeColorSel    =   16777215
      BackColorBkg    =   12632256
      SelectionMode   =   1
   End
End
Attribute VB_Name = "consulta_tabela_provento_desconto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim lFlagConsulta As Integer
    Dim lSQl As String
    Dim fld As adodb.Field
    Dim lRst As adodb.Recordset
Private Sub AtualizaGrid()
    On Error GoTo ErroConsulta
    Dim xOperando As String
    Dim xCondicao As String
    Dim xNomeCampo As String
    xNomeCampo = cbo_campo.Text
    If Mid(xNomeCampo, 1, 1) = "[" Then
        xNomeCampo = Mid(xNomeCampo, 2, Len(xNomeCampo) - 2)
    End If
    xCondicao = txt_condicao
    For Each fld In lRst.Fields
        If fld.Name = xNomeCampo Then
        '10 = 202
            If fld.Type = 8 Then
                xCondicao = "#" & Format(CDate(xCondicao), "yyyy/mm/dd") & "#"
            ElseIf (fld.Type = 200 Or fld.Type = 202) And cbo_operador.Text = "Semelhante" Then
                xCondicao = Chr(39) & "%" & xCondicao & "%" & Chr(39)
            ElseIf (fld.Type = 200 Or fld.Type = 202) And cbo_operador.Text <> "Semelhante" Then
                xCondicao = Chr(34) & xCondicao & Chr(34)
            End If
            Exit For
        End If
    Next
    If cbo_operador.Text = "Diferente" Then
        xOperando = "<>"
    ElseIf cbo_operador.Text = "Igual" Then
        xOperando = "="
    ElseIf cbo_operador.Text = "Maior" Then
        xOperando = ">"
    ElseIf cbo_operador.Text = "Maior Igual" Then
        xOperando = ">="
    ElseIf cbo_operador.Text = "Menor" Then
        xOperando = "<"
    ElseIf cbo_operador.Text = "Menor Igual" Then
        xOperando = "<="
    ElseIf cbo_operador.Text = "Semelhante" Then
        xOperando = "Like"
    End If
    If ValidaCampos Then
        'Prepara SQL
        lSQl = ""
        lSQl = lSQl & "SELECT Codigo, Nome, Percentual, Valor, Fracao, [Provento ou Desconto], Automatico,"
        lSQl = lSQl & "       [Base para Calculo]"
        lSQl = lSQl & "  FROM Tabela_Provento_Desconto"
        If cbo_campo.ListIndex <> -1 Then
            lSQl = lSQl & " WHERE Tabela_Provento_Desconto." & cbo_campo.Text & " " & xOperando & " " & xCondicao
        End If
        lSQl = lSQl & " ORDER BY Nome"
        'Abre RecordSet
        Set lRst = Nothing
        Set lRst = New adodb.Recordset
        Set lRst = Conectar.RsConexao(lSQl)
        'as proximas 2 linhas substitui a linha anterior
        'lRst.CursorLocation = adUseClient
        'lRst.Open lSQl, cnnSGP, adOpenForwardOnly, adLockReadOnly
        FormaGrid
        MontaGrid
    End If
    Exit Sub
ErroConsulta:
    If Err = 3075 Then
        MsgBox "Condição inválida.", vbInformation, "Erro de Consulta"
    Else
        MsgBox Error, vbInformation, "Erro de Consulta"
    End If
    Exit Sub
End Sub
Private Sub Finaliza()
    Set lRst = Nothing
End Sub
Private Sub FormaGrid()
    Dim i As Integer
    MSFlexGrid.WordWrap = True
    MSFlexGrid.Cols = 8
    MSFlexGrid.RowHeight(0) = 500
    MSFlexGrid.Rows = 1
    i = 0
    MSFlexGrid.TextMatrix(0, i) = "Código"
    MSFlexGrid.ColAlignment(i) = flexAlignGeneral
    MSFlexGrid.ColWidth(i) = 600
    i = i + 1
    MSFlexGrid.TextMatrix(0, i) = "Descrição"
    MSFlexGrid.ColAlignment(i) = flexAlignGeneral
    MSFlexGrid.ColWidth(i) = 2500
    i = i + 1
    MSFlexGrid.TextMatrix(0, i) = "Percentual"
    MSFlexGrid.ColAlignment(i) = flexAlignGeneral
    MSFlexGrid.ColWidth(i) = 1000
    i = i + 1
    MSFlexGrid.TextMatrix(0, i) = "Valor"
    MSFlexGrid.ColAlignment(i) = flexAlignGeneral
    MSFlexGrid.ColWidth(i) = 1000
    i = i + 1
    MSFlexGrid.TextMatrix(0, i) = "Fração"
    MSFlexGrid.ColAlignment(i) = flexAlignGeneral
    MSFlexGrid.ColWidth(i) = 1000
    i = i + 1
    MSFlexGrid.TextMatrix(0, i) = "Prov. Desc."
    MSFlexGrid.ColAlignment(i) = flexAlignGeneral
    MSFlexGrid.ColWidth(i) = 700
    i = i + 1
    MSFlexGrid.TextMatrix(0, i) = "Automático"
    MSFlexGrid.ColAlignment(i) = flexAlignGeneral
    MSFlexGrid.ColWidth(i) = 900
    i = i + 1
    MSFlexGrid.TextMatrix(0, i) = "Base para Cálculo"
    MSFlexGrid.ColAlignment(i) = flexAlignGeneral
    MSFlexGrid.ColWidth(i) = 3000
End Sub
Private Sub MarcaCelulas()
    g_string = ""
    If Val(MSFlexGrid.TextMatrix(MSFlexGrid.Row, 0)) > 0 Then
        g_string = MSFlexGrid.TextMatrix(MSFlexGrid.Row, 0) & "|@|"
    End If
End Sub
Private Sub MontaGrid()
    Dim i As Integer
    Dim i2 As Integer
    If Not lRst.EOF Then
        MSFlexGrid.Visible = False
        lRst.MoveFirst
        Do Until lRst.EOF
            MSFlexGrid.Rows = MSFlexGrid.Rows + 1
            MSFlexGrid.Row = MSFlexGrid.Rows - 1
            i2 = -1
            For Each fld In lRst.Fields
                i2 = i2 + 1
                MSFlexGrid.Col = i2
                If fld.Name = "Automatico" Then
                    If fld.Value = True Then
                        MSFlexGrid.Text = "Sim"
                    Else
                        MSFlexGrid.Text = "Não"
                    End If
                ElseIf fld.Name = "Percentual" Then
                    MSFlexGrid.Text = Format(fld.Value, "##0.00") & "%"
                ElseIf fld.Name = "Valor" Then
                    MSFlexGrid.Text = Format(fld.Value, "###,###,##0.00")
                Else
                    If IsNull(fld.Value) Then
                        MSFlexGrid.Text = ""
                    Else
                        MSFlexGrid.Text = fld.Value
                    End If
                End If
            Next
            lRst.MoveNext
        Loop
        MSFlexGrid.Visible = True
        MSFlexGrid.Row = 1
        MSFlexGrid.Col = 0
    End If
End Sub
Private Sub PreencheCampos()
    Dim i As Integer
    Dim x_string As String
    For Each fld In lRst.Fields
        x_string = fld.Name
        For i = 1 To Len(fld.Name)
            If Mid(fld.Name, i, 1) = " " Then
                x_string = "[" & fld.Name & "]"
                Exit For
            End If
        Next
        cbo_campo.AddItem x_string
    Next
End Sub
Private Sub PreencheOperador()
    cbo_operador.Clear
    cbo_operador.AddItem "Diferente"
    cbo_operador.AddItem "Igual"
    cbo_operador.AddItem "Maior"
    cbo_operador.AddItem "Maior Igual"
    cbo_operador.AddItem "Menor"
    cbo_operador.AddItem "Menor Igual"
    cbo_operador.AddItem "Semelhante"
End Sub
Private Sub cbo_campo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_operador.SetFocus
    End If
End Sub
Private Sub cbo_operador_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_condicao.SetFocus
    End If
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If cbo_campo.ListIndex = -1 Then
        MsgBox "Informe o campo a ser testado.", vbInformation, "Atenção!"
        cbo_campo.SetFocus
    ElseIf cbo_operador.ListIndex = -1 Then
        MsgBox "Informe o operando a ser testado.", vbInformation, "Atenção!"
        cbo_operador.SetFocus
    ElseIf txt_condicao = "" Then
        MsgBox "Informe a condição testada.", vbInformation, "Atenção!"
        txt_condicao.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_ok_Click()
    Call GravaConfiguracaoConsulta(Me.Name, cbo_campo.Text, cbo_operador.Text, txt_condicao.Text)
    AtualizaGrid
    MSFlexGrid.SetFocus
End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub Form_Activate()
    If lFlagConsulta = 0 Then
        FormaGrid
        Inicializa
        txt_condicao.SetFocus
    Else
        lFlagConsulta = 0
    End If
End Sub
Private Sub Inicializa()
    Dim i As Integer
    Dim x_campo As String
    Dim x_operador As String
    x_campo = "Nome"
    x_operador = "Semelhante"
    txt_condicao = "A"
    If BuscaConfiguracaoConsulta(Me.Name) Then
        x_campo = RetiraGString(1)
        x_operador = RetiraGString(2)
        txt_condicao = RetiraGString(3)
        g_string = ""
    End If
    cbo_campo.ListIndex = -1
    For i = 0 To cbo_campo.ListCount - 1
        cbo_campo.ListIndex = i
        If cbo_campo = x_campo Then
            Exit For
        End If
    Next
    cbo_operador.ListIndex = -1
    For i = 0 To cbo_operador.ListCount - 1
        cbo_operador.ListIndex = i
        If cbo_operador = x_operador Then
            Exit For
        End If
    Next
    AtualizaGrid
End Sub
Private Sub Form_Deactivate()
    lFlagConsulta = 1
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        cmd_sair_Click
    ElseIf KeyCode = vbKeyF11 Then
        KeyCode = 0
        cmd_ok_Click
    End If
End Sub
Private Sub Form_Load()
    Screen.MousePointer = 1
    CentraForm Me
    lSQl = "SELECT * FROM Tabela_Provento_Desconto WHERE Codigo = 0"
    Set lRst = New adodb.Recordset
    Set lRst = Conectar.RsConexao(lSQl)
    PreencheCampos
    lRst.Close
    PreencheOperador
    g_string = ""
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub MSFlexGrid_DblClick()
    MarcaCelulas
    cmd_sair_Click
End Sub
Private Sub MSFlexGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        MarcaCelulas
        cmd_sair_Click
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
        MarcaCelulas
        cmd_sair_Click
    End If
End Sub
Private Sub txt_condicao_GotFocus()
    txt_condicao.SelStart = 0
    txt_condicao.SelLength = Len(txt_condicao)
End Sub
Private Sub txt_condicao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok_Click
    End If
End Sub
