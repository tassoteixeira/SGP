VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form baixa_cheque 
   Caption         =   "Baixa de Cheques"
   ClientHeight    =   2085
   ClientLeft      =   1920
   ClientTop       =   2790
   ClientWidth     =   4830
   Icon            =   "baixa_cheque.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "baixa_cheque.frx":030A
   ScaleHeight     =   2085
   ScaleWidth      =   4830
   Begin VB.OptionButton opt_vencimento 
      Caption         =   "Vencimento"
      Height          =   255
      Left            =   3060
      TabIndex        =   8
      Top             =   780
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.OptionButton opt_emissao 
      Caption         =   "Emissao"
      Height          =   255
      Left            =   1680
      TabIndex        =   7
      Top             =   780
      Width           =   1215
   End
   Begin VB.CommandButton cmd_data_f 
      Height          =   315
      Left            =   4200
      Picture         =   "baixa_cheque.frx":0750
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Selecione a data pelo calendário."
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton cmd_data_i 
      Height          =   315
      Left            =   1260
      Picture         =   "baixa_cheque.frx":1A2A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Selecione a data pelo calendário."
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton cmd_baixa 
      Caption         =   "&Baixar"
      Height          =   855
      Left            =   120
      Picture         =   "baixa_cheque.frx":2D04
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Baixa os cheques no período informado."
      Top             =   1140
      Width           =   795
   End
   Begin VB.CommandButton cmd_estornar 
      Caption         =   "&Estornar"
      Height          =   855
      Left            =   2040
      Picture         =   "baixa_cheque.frx":3FDE
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Estorna os cheques baixados no período informado."
      Top             =   1140
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3900
      Picture         =   "baixa_cheque.frx":52B8
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   1140
      Width           =   795
   End
   Begin MSMask.MaskEdBox msk_data_final 
      Height          =   300
      Left            =   3060
      TabIndex        =   4
      Top             =   360
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
      TabIndex        =   1
      Top             =   360
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label7 
      Caption         =   "&Periodo por"
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   780
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "D&ata final"
      Height          =   255
      Index           =   2
      Left            =   3060
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "&Data inicial"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "baixa_cheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lSQL As String
Dim lCampo As String
Dim lTotal As Currency

Dim rst_cheque As New adodb.Recordset
Dim rst_baixa_cheque As New adodb.Recordset
Dim rstTotal As New adodb.Recordset

Private Sub PreparaBaixa()
    Dim xString As String
    
    lSQL = "SELECT SUM(Valor) AS Total"
    lSQL = lSQL & " FROM Movimento_Cheque"
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    lSQL = lSQL & " AND " & lCampo & " >= " & preparaData(CDate(msk_data_inicial.Text))
    lSQL = lSQL & " AND " & lCampo & " <= " & preparaData(CDate(msk_data_final.Text))
    Set rstTotal = Conectar.RsConexao(lSQL)
    If rstTotal.RecordCount > 0 Then
        If Not IsNull(rstTotal!total) Then
            lTotal = rstTotal!total
            If (MsgBox("No período informado tem R$ " & Format(lTotal, "###,###,##0.00") & " em cheques à serem baixados." & Chr(13) & Chr(10) & Chr(10) & "Deseja realmente baixá-los?", vbQuestion + vbYesNo + vbDefaultButton2, "Baixa de Cheques.")) = vbYes Then
                If opt_emissao.Value = True Then
                    xString = "Data Emissão:"
                Else
                    xString = "Data Vencimento:"
                End If
                xString = xString & msk_data_inicial.Text & " a " & msk_data_final.Text
                Call GravaAuditoria(1, Me.name, 18, xString)
                Baixa
            End If
        Else
            MsgBox "Não existe cheque à ser baixado no período informado!", 48, "Baixa de Cheques."
        End If
    End If
    rstTotal.Close
    AtivaBotoes (True)
    cmd_sair.SetFocus
End Sub
Private Sub PreparaEstorno()
    Dim xString As String
    
    lSQL = "SELECT SUM(Valor) AS Total"
    lSQL = lSQL & " FROM Baixa_Cheque"
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    lSQL = lSQL & " AND " & lCampo & " >= " & preparaData(CDate(msk_data_inicial.Text))
    lSQL = lSQL & " AND " & lCampo & " <= " & preparaData(CDate(msk_data_final.Text))
    Set rstTotal = Conectar.RsConexao(lSQL)
    If rstTotal.RecordCount > 0 Then
        If Not IsNull(rstTotal!total) Then
            lTotal = rstTotal!total
            If (MsgBox("No período informado tem R$ " & Format(lTotal, "###,###,##0.00") & " em cheques baixados à serem estornados." & Chr(13) & Chr(10) & Chr(10) & "Deseja realmente estorná-los?", vbQuestion + vbYesNo + vbDefaultButton2, "Estorno de Cheque Baixado.")) = vbYes Then
                If opt_emissao.Value = True Then
                    xString = "Data Emissão:"
                Else
                    xString = "Data Vencimento:"
                End If
                xString = xString & msk_data_inicial.Text & " a " & msk_data_final.Text
                Call GravaAuditoria(1, Me.name, 19, xString)
                Estorno
            End If
        Else
            MsgBox "Não existe cheque baixado à ser estornado no período informado!", vbExclamation, "Estorno de Cheque Baixado."
        End If
    End If
    rstTotal.Close
    AtivaBotoes (True)
    cmd_sair.SetFocus
End Sub
Private Sub AtivaBotoes(xAtiva As Boolean)
    cmd_baixa.Enabled = xAtiva
    cmd_estornar.Enabled = xAtiva
    cmd_sair.Enabled = xAtiva
End Sub
Private Sub Baixa()
    Dim xMensagem As String
    On Error GoTo FileError
    
    'INCLUI BAIXA_CHEQUE
    lSQL = "SELECT *"
    lSQL = lSQL & " FROM Movimento_Cheque"
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    lSQL = lSQL & " AND " & lCampo & " >= " & preparaData(CDate(msk_data_inicial.Text))
    lSQL = lSQL & " AND " & lCampo & " <= " & preparaData(CDate(msk_data_final.Text))
    Set rst_cheque = Conectar.RsConexao(lSQL)
    Conectar.IniciaTransacao
    If rst_cheque.RecordCount > 0 Then
        rst_cheque.MoveFirst
        Do Until rst_cheque.EOF
            Call GravaBaixa
            rst_cheque.MoveNext
        Loop
    End If
    rst_cheque.Close
    
    'DELETA MOVIMENTO_CHEQUE
    lSQL = "DELETE"
    lSQL = lSQL & " FROM Movimento_Cheque"
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    lSQL = lSQL & " AND " & lCampo & " >= " & preparaData(CDate(msk_data_inicial.Text))
    lSQL = lSQL & " AND " & lCampo & " <= " & preparaData(CDate(msk_data_final.Text))
    If Conectar.ExecutaSql(lSQL) > 0 Then
        Conectar.ConfirmaTransacao
        xMensagem = "Baixa de cheques concluida!"
        MsgBox xMensagem, vbExclamation, "Baixa de Cheque."
        Call GravaAuditoria(1, Me.name, 22, xMensagem & " Total:" & Format(lTotal, "###,###,##0.00"))
        
    Else
        Conectar.CancelaTransacao
        xMensagem = "Não foi possível excluir Movimento_Cheque!"
        MsgBox xMensagem, vbCritical, "Erro de Integridade."
        Call GravaAuditoria(1, Me.name, 22, xMensagem & " Total:" & Format(lTotal, "###,###,##0.00"))
    End If
    Exit Sub
FileError:
    Conectar.CancelaTransacao
    'ErroArquivo tbl_baixa_cheque.Name, "Cheque Baixadoo"
    MsgBox "Erro Baixa: " & Error
    Exit Sub
End Sub
Private Sub Estorno()
    Dim xMensagem As String
    On Error GoTo FileError
    
    'INCLUI MOVIMENTO_CHEQUE
    lSQL = "SELECT *"
    lSQL = lSQL & " FROM Baixa_Cheque"
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    lSQL = lSQL & " AND " & lCampo & " >= " & preparaData(CDate(msk_data_inicial.Text))
    lSQL = lSQL & " AND " & lCampo & " <= " & preparaData(CDate(msk_data_final.Text))
    Set rst_baixa_cheque = Conectar.RsConexao(lSQL)
    Conectar.IniciaTransacao
    If rst_baixa_cheque.RecordCount > 0 Then
        rst_baixa_cheque.MoveFirst
        Do Until rst_baixa_cheque.EOF
            gSQL = "INSERT INTO Movimento_Cheque ( empresa, [data de emissao], [numero da conta], [numero do cheque], periodo, "
            gSQL = gSQL & "[tipo do movimento], valor, [data do vencimento], emitente, [ordem da digitacao], "
            gSQL = gSQL & "[codigo de barra 1], [codigo de barra 2], [codigo de barra 3], [banco agencia], "
            gSQL = gSQL & "Telefone, [Numero do Movimento do Caixa], [Codigo do Vendedor], [CPF CNPJ], "
            gSQL = gSQL & "[Numero da Ilha], [Data da Custodia] ) VALUES ( "
            Call sqlNumero(1, rst_baixa_cheque!Empresa, ", ")
            Call sqlData(1, rst_baixa_cheque![Data de Emissao], ", ")
            Call sqlTexto(1, rst_baixa_cheque![Numero da Conta], ", ")
            Call sqlTexto(1, rst_baixa_cheque![Numero do Cheque], ", ")
            Call sqlTexto(1, rst_baixa_cheque!Periodo, ", ")
            Call sqlTexto(1, rst_baixa_cheque![Tipo do Movimento], ", ")
            Call sqlValor(1, rst_baixa_cheque!Valor, ", ")
            Call sqlData(1, rst_baixa_cheque![Data do Vencimento], ", ")
            Call sqlTexto(1, rst_baixa_cheque!Emitente, ", ")
            Call sqlNumero(1, rst_baixa_cheque![Ordem da Digitacao], ", ")
            Call sqlTexto(1, rst_baixa_cheque![Codigo de Barra 1], ", ")
            Call sqlTexto(1, rst_baixa_cheque![Codigo de Barra 2], ", ")
            Call sqlTexto(1, rst_baixa_cheque![Codigo de Barra 3], ", ")
            Call sqlTexto(1, rst_baixa_cheque![Banco Agencia], ", ")
            Call sqlTexto(1, rst_baixa_cheque!Telefone, ", ")
            Call sqlNumero(1, rst_baixa_cheque![Numero do Movimento do Caixa], ", ")
            Call sqlNumero(1, rst_baixa_cheque![Codigo do Vendedor], ", ")
            Call sqlTexto(1, rst_baixa_cheque![CPF CNPJ], ", ")
            Call sqlNumero(1, rst_baixa_cheque![Numero da Ilha], ", ")
            If IsDate(rst_baixa_cheque![Data da Custodia]) Then
                Call sqlData(1, rst_baixa_cheque![Data da Custodia], " )")
            Else
                Call sqlData(1, "00:00:00", " )")
            End If
            If Conectar.ExecutaSql(gSQL) = 0 Then
                xMensagem = "Não foi possível incluir Movimento_Cheque!"
                MsgBox xMensagem, vbCritical, "Erro de Integridade."
                Call GravaAuditoria(1, Me.name, 22, xMensagem & " Ch:" & rst_baixa_cheque![Numero do Cheque])
            End If
            rst_baixa_cheque.MoveNext
        Loop
    End If
    rst_baixa_cheque.Close
    
    'DELETA BAIXA_CHEQUE
    lSQL = "DELETE"
    lSQL = lSQL & " FROM Baixa_Cheque"
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    lSQL = lSQL & " AND " & lCampo & " >= " & preparaData(CDate(msk_data_inicial.Text))
    lSQL = lSQL & " AND " & lCampo & " <= " & preparaData(CDate(msk_data_final.Text))
    If Conectar.ExecutaSql(lSQL) > 0 Then
        Conectar.ConfirmaTransacao
        xMensagem = "Estorno dos cheques baixados concluido!"
        MsgBox xMensagem, vbExclamation, "Estorno de Cheque Baixado."
        Call GravaAuditoria(1, Me.name, 22, xMensagem & " Total:" & Format(lTotal, "###,###,##0.00"))
    Else
        Conectar.CancelaTransacao
        xMensagem = "Não foi possível estornar Baixa_Cheque!"
        MsgBox xMensagem, vbCritical, "Erro de Integridade."
        Call GravaAuditoria(1, Me.name, 22, xMensagem & " Total:" & Format(lTotal, "###,###,##0.00"))
    End If
    Exit Sub
FileError:
    Conectar.CancelaTransacao
    MsgBox "Erro Estorno: " & Error
    Exit Sub
End Sub
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
End Sub
Function GravaBaixa() As Boolean
    Dim xMensagem As String
    On Error GoTo FileError
    
    GravaBaixa = False
    'Grava Baixa_Cheque
    gSQL = "INSERT INTO Baixa_Cheque ( empresa, [data de emissao], [numero da conta], [numero do cheque], periodo, "
    gSQL = gSQL & "[tipo do movimento], agencia, valor, [data do vencimento], emitente, [ordem da digitacao], "
    gSQL = gSQL & "[codigo de barra 1], [codigo de barra 2], [codigo de barra 3], [banco agencia], Telefone, "
    gSQL = gSQL & "[Numero do Movimento do Caixa], [Codigo do Vendedor], [CPF CNPJ], [Numero da Ilha], "
    gSQL = gSQL & "[Data da Custodia], [Data do Pagamento], [Periodo do Pagamento] ) VALUES ( "
    Call sqlNumero(1, rst_cheque!Empresa, ", ")
    Call sqlData(1, rst_cheque![Data de Emissao], ", ")
    Call sqlTexto(1, rst_cheque![Numero da Conta], ", ")
    Call sqlTexto(1, rst_cheque![Numero do Cheque], ", ")
    Call sqlTexto(1, rst_cheque!Periodo, ", ")
    Call sqlTexto(1, rst_cheque![Tipo do Movimento], ", ")
    Call sqlTexto(1, Mid(rst_cheque![Banco Agencia], 4, 4), ", ")
    Call sqlValor(1, rst_cheque!Valor, ", ")
    Call sqlData(1, rst_cheque![Data do Vencimento], ", ")
    Call sqlTexto(1, rst_cheque!Emitente, ", ")
    Call sqlNumero(1, rst_cheque![Ordem da Digitacao], ", ")
    Call sqlTexto(1, rst_cheque![Codigo de Barra 1], ", ")
    Call sqlTexto(1, rst_cheque![Codigo de Barra 2], ", ")
    Call sqlTexto(1, rst_cheque![Codigo de Barra 3], ", ")
    Call sqlTexto(1, rst_cheque![Banco Agencia], ", ")
    Call sqlTexto(1, rst_cheque!Telefone, ", ")
    Call sqlNumero(1, rst_cheque![Numero do Movimento do Caixa], ", ")
    Call sqlNumero(1, rst_cheque![Codigo do Vendedor], ", ")
    Call sqlTexto(1, rst_cheque![CPF CNPJ], ", ")
    Call sqlNumero(1, rst_cheque![Numero da Ilha], ", ")
    If IsDate(rst_cheque![Data da Custodia]) Then
        Call sqlData(1, rst_cheque![Data da Custodia], ", ")
    Else
        Call sqlData(1, "00:00:00", ", ")
    End If
    'If IsDate(rst_cheque![Data do Pagamento]) Then
        Call sqlData(1, Date, ", ")
    'Else
    '    Call sqlData(1, "00:00:00", ", ")
    'End If
    Call sqlNumero(1, 1, " )")
    
    If Conectar.ExecutaSql(gSQL) = 0 Then
        xMensagem = "Não foi possível incluir Baixa_Cheque!"
        MsgBox xMensagem, vbCritical, "Erro de Integridade."
        Call GravaAuditoria(1, Me.name, 22, xMensagem & " Ch:" & rst_cheque![Numero do Cheque])
    End If
    GravaBaixa = True
    Exit Function

FileError:
    MsgBox "Baixa Existente" & Chr(10) & Error
    Exit Function
End Function
Private Sub cmd_baixa_Click()
    If opt_vencimento.Value = True Then
        lCampo = "[Data do Vencimento]"
    Else
        lCampo = "[Data de Emissao]"
    End If
    If ValidaCampos Then
        AtivaBotoes (False)
        PreparaBaixa
    End If
End Sub
Private Sub cmd_data_f_Click()
    g_string = msk_data_final.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_inicial.Text = RetiraGString(1)
        msk_data_final.Text = RetiraGString(2)
    Else
        msk_data_final.Text = RetiraGString(1)
    End If
    g_string = " "
    cmd_baixa.SetFocus
End Sub
Private Sub cmd_data_i_Click()
    g_string = msk_data_inicial
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_inicial.Text = RetiraGString(1)
        msk_data_final.Text = RetiraGString(2)
        cmd_baixa.SetFocus
    Else
        msk_data_inicial.Text = RetiraGString(1)
        msk_data_final.SetFocus
    End If
    g_string = " "
End Sub
Private Sub cmd_estornar_Click()
    If opt_vencimento.Value = True Then
        lCampo = "[Data do Vencimento]"
    Else
        lCampo = "[Data de Emissao]"
    End If
    If ValidaCampos Then
        AtivaBotoes (False)
        PreparaEstorno
    End If
End Sub
Function ValidaCampos() As Boolean
    ValidaCampos = False
    If Not IsDate(msk_data_inicial.Text) Then
        MsgBox "Informe a data da inicial.", vbInformation, "Atenção!"
        msk_data_inicial.SetFocus
    ElseIf Not IsDate(msk_data_final.Text) Then
        MsgBox "Informe a data da final.", vbInformation, "Atenção!"
        msk_data_final.SetFocus
    ElseIf CDate(msk_data_final.Text) < CDate(msk_data_inicial.Text) Then
        MsgBox "A data final deve ser maior ou igual a " & msk_data_inicial.Text & ".", vbInformation, "Atenção!"
        msk_data_final.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    BuscaDatas
    cmd_baixa.SetFocus
    Screen.MousePointer = 1
End Sub
Private Sub BuscaDatas()
    msk_data_inicial.Text = Format(CDate(g_data_def - 1), "dd/mm/yyyy")
    msk_data_final.Text = Format(CDate(g_data_def - 1), "dd/mm/yyyy")
    If opt_vencimento.Value = True Then
        lCampo = "[Data do Vencimento]"
    Else
        lCampo = "[Data de Emissao]"
    End If
    lSQL = "SELECT TOP 1 " & lCampo
    lSQL = lSQL & " FROM Movimento_Cheque"
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    lSQL = lSQL & " ORDER BY " & lCampo
    Set rst_cheque = Conectar.RsConexao(lSQL)
    If rst_cheque.RecordCount > 0 Then
        If opt_vencimento.Value = True Then
            msk_data_inicial.Text = Format(rst_cheque![Data do Vencimento], "dd/mm/yyyy")
        Else
            msk_data_inicial.Text = Format(rst_cheque![Data de Emissao], "dd/mm/yyyy")
        End If
    End If
    rst_cheque.Close
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        cmd_sair_Click
    End If
End Sub
Private Sub Form_Load()
    CentraForm Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub Label7_Click()
    BuscaDatas
End Sub
Private Sub msk_data_final_GotFocus()
    msk_data_final.SelStart = 0
    msk_data_final.SelLength = 5
End Sub
Private Sub msk_data_final_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_baixa.SetFocus
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
Private Sub opt_emissao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_baixa.SetFocus
    End If
End Sub
Private Sub opt_vencimento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_baixa.SetFocus
    End If
End Sub
