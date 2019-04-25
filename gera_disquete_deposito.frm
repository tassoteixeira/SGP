VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form gera_disquete_deposito 
   Caption         =   "Gera Disquete para Depósito Bancário"
   ClientHeight    =   2835
   ClientLeft      =   1920
   ClientTop       =   2790
   ClientWidth     =   4830
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "gera_disquete_deposito.frx":0000
   ScaleHeight     =   2835
   ScaleWidth      =   4830
   Begin VB.Frame frm_dados 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   4575
      Begin VB.TextBox txt_remessa 
         Height          =   315
         Left            =   120
         MaxLength       =   4
         TabIndex        =   6
         Top             =   1200
         Width           =   675
      End
      Begin VB.TextBox txt_bordero 
         Height          =   315
         Left            =   3360
         MaxLength       =   3
         TabIndex        =   8
         Top             =   1200
         Width           =   675
      End
      Begin MSMask.MaskEdBox msk_data_final 
         Height          =   300
         Left            =   3360
         TabIndex        =   4
         Top             =   480
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
         TabIndex        =   2
         Top             =   480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "&Data inicial"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "D&ata final"
         Height          =   255
         Index           =   2
         Left            =   3360
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "&Remessa"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "&Borderô"
         Height          =   255
         Left            =   3360
         TabIndex        =   7
         Top             =   960
         Width           =   735
      End
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   1140
      Picture         =   "gera_disquete_deposito.frx":0446
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Confirma a geração do disquete para depósito."
      Top             =   1860
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Caption         =   "&Sair"
      Height          =   855
      Left            =   2880
      Picture         =   "gera_disquete_deposito.frx":1720
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   1860
      Width           =   795
   End
End
Attribute VB_Name = "gera_disquete_deposito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_bordero As Long
Dim tbl_movimento_cheque_avista As Table
Dim lSQl As String

Private MovCheque As New cMovimentoCheque
Private rsCheque As New adodb.Recordset
Private Sub CriaDisqueteChequeAVista()
    Dim x_total As Currency
    Dim i As Integer
    x_total = TotalChequeAVista(0)
    If x_total > 0 Then
        If (MsgBox("Insira o disquete de " & Chr(34) & "CHEQUE À VISTA" & Chr(34) & " no drive " & Chr(34) & "A:" & Chr(34) & " e clique em OK.", vbOKCancel + vbDefaultButton1, "Geração de Disquete - À VISTA")) = 1 Then
            Open "A:CHEQUES.DAT" For Output As #1
            Call FazCabecalhoAVista(0)
            For i = 1 To 10
                If TotalChequeAVista(i) > 0 Then
                    Call FazDetalheAVista(i)
                End If
            Next
            MsgBox "Operação concluída!", 48, "Geração de Disquete - À VISTA"
            Close #1
            cmd_sair.SetFocus
        End If
    Else
        MsgBox "Não existe cheques pré-datados no período informado!", 48, "Geração de Disquete - CUSTÓDIA"
        cmd_sair.SetFocus
    End If
End Sub
Private Sub CriaDisqueteChequePreDatado()
    Dim x_total As Currency
    Dim i As Integer
    x_total = MovCheque.TotalEmissaoPeriodo(g_empresa, CDate(msk_data_inicial.Text), CDate(msk_data_final.Text), "1", "9", "0", "P")
    If x_total > 0 Then
        If (MsgBox("Insira o disquete de " & Chr(34) & "CHEQUE CUSTÓDIA" & Chr(34) & " no drive " & Chr(34) & "A:" & Chr(34) & " e clique em OK.", vbOKCancel + vbDefaultButton1, "Geração de Disquete - CUSTÓDIA")) = 1 Then
            Open "A:CUSTODIA.DAT" For Output As #1
            l_bordero = CLng(txt_bordero) - 1
            For i = 1 To 10
                If x_total > 0 Then
                    Call FazCabecalhoPreDatado(i)
                    Call FazDetalhePreDatado(i)
                End If
            Next
            MsgBox "Operação concluída!", 48, "Geração de Disquete - CUSTÓDIA"
            Close #1
            cmd_sair.SetFocus
        End If
    Else
        MsgBox "Não existe cheques pré-datados no período informado!", 48, "Geração de Disquete - CUSTÓDIA"
        cmd_sair.SetFocus
    End If
End Sub
Private Sub FazCabecalhoAVista(x_empresa As Integer)
    Dim x_dados As String
    Dim x_total As Currency
    x_total = TotalChequeAVista(x_empresa)
    'Zero Fixo
    x_dados = Space(48)
    Mid(x_dados, 1, 1) = "0"
    'Código do Bradesco
    Mid(x_dados, 2, 3) = "237"
    'Agência
    Mid(x_dados, 5, 4) = "0638"
    'Zero Fixo
    Mid(x_dados, 9, 1) = "0"
    'Nove Fixos
    Mid(x_dados, 10, 10) = "9999999999"
    'Zeros Fixos
    Mid(x_dados, 20, 8) = "00000000"
    'Empresa
    Mid(x_dados, 28, 3) = Format(2, "000")
    'Zeros Fixos
    Mid(x_dados, 31, 5) = "00000"
    'Total do Depósito
    Mid(x_dados, 36, 13) = Mid(Format(x_total, "00000000000.00"), 1, 11) & Mid(Format(x_total, "00000000000.00"), 13, 2)
    Print #1, x_dados
End Sub
Private Sub FazDetalhePreDatado(x_empresa As Integer)
    Dim x_dados As String
    
    'Verifica Movimento_Cheque
    'Prepara SQL
    lSQl = ""
    lSQl = lSQl & "SELECT [Data do Vencimento], [Data de Emissao], Valor, [Codigo de Barra 1], [Codigo de Barra 2], [Codigo de Barra 3]"
    lSQl = lSQl & "  FROM Movimento_Cheque"
    lSQl = lSQl & " WHERE Empresa = " & g_empresa
    lSQl = lSQl & "   AND [Data de Emissao] >= " & preparaData(msk_data_inicial.Text)
    lSQl = lSQl & "   AND [Data de Emissao] <= " & preparaData(msk_data_final.Text)
    lSQl = lSQl & " ORDER BY [Data do Vencimento], [Data de Emissao], Periodo, [Tipo do Movimento], [Ordem da Digitacao], [Numero da Conta], [Numero do Cheque]"
    'Abre RecordSet
    Set rsCheque = New adodb.Recordset
    Set rsCheque = Conectar.RsConexao(lSQl)
    If rsCheque.RecordCount > 0 Then
        Do Until rsCheque.EOF
            'Zeros Fixos
            x_dados = Space(120)
            Mid(x_dados, 1, 11) = "00000000000"
            'Data de Vencimento
            Mid(x_dados, 12, 6) = Format(rsCheque("Data do Vencimento").Value, "ddmmyy")
            'Codigo de Barra 2
            Mid(x_dados, 18, 10) = rsCheque("Codigo de Barra 2").Value
            '(1,1) Codigo do Codigo de Barra 3
            Mid(x_dados, 28, 1) = Mid(rsCheque("Codigo de Barra 3").Value, 1, 1)
            'Codigo de Barra 1
            Mid(x_dados, 29, 8) = rsCheque("Codigo de Barra 2").Value
            '(2,11) Codigo do Codigo de Barra 3
            Mid(x_dados, 37, 11) = Mid(rsCheque("Codigo de Barra 2").Value, 2, 11)
            'Valor do Cheque
            Mid(x_dados, 48, 13) = Mid(Format(rsCheque("Valor").Value, "00000000000.00"), 1, 11) & Mid(Format(rsCheque("Valor").Value, "00000000000.00"), 13, 2)
            'Letra "D" Fixo
            Mid(x_dados, 61, 1) = "D"
            Print #1, x_dados
            rsCheque.MoveNext
        Loop
    End If
    If rsCheque.State = 1 Then
        rsCheque.Close
    End If
End Sub
Private Sub FazCabecalhoPreDatado(x_empresa As Integer)
    Dim x_dados As String
    Dim x_total As Currency
    x_total = MovCheque.TotalEmissaoPeriodo(x_empresa, CDate(msk_data_inicial.Text), CDate(msk_data_final.Text), "1", "9", "0", "V")
    x_dados = Space(120)
    'Agência
    Mid(x_dados, 1, 4) = "0638"
    'Conta
    If x_empresa = 2 Then
        Mid(x_dados, 5, 7) = "0055948"
    ElseIf x_empresa = 3 Then
        Mid(x_dados, 5, 7) = "0055956"
    ElseIf x_empresa = 4 Then
        Mid(x_dados, 5, 7) = "0055980"
    ElseIf x_empresa = 6 Then
        Mid(x_dados, 5, 7) = "0055964"
    ElseIf x_empresa = 9 Then
        Mid(x_dados, 5, 7) = "0055972"
    ElseIf x_empresa = 10 Then
        Mid(x_dados, 5, 7) = "0055778"
    ElseIf x_empresa = 11 Then
        Mid(x_dados, 5, 7) = "0012602"
    Else
        Mid(x_dados, 5, 7) = "0051349"
    End If
    'Zeros Fixos
    Mid(x_dados, 12, 2) = "00"
    'Lote
    Mid(x_dados, 14, 3) = Format(x_empresa, "000")
    'Total do Depósito
    Mid(x_dados, 17, 18) = "00000" & Mid(Format(x_total, "00000000000.00"), 1, 11) & Mid(Format(x_total, "00000000000.00"), 13, 2)
    'Código do Bradesco
    Mid(x_dados, 35, 3) = "237"
    'Zeros Fixos
    Mid(x_dados, 38, 4) = "0000"
    'Lote
    Mid(x_dados, 42, 3) = Format(x_empresa, "000")
    'Zeros Fixos
    Mid(x_dados, 45, 2) = "00"
    'Polo
    Mid(x_dados, 47, 4) = "4436"
    'Data do Deposito
    Mid(x_dados, 51, 6) = Format(CDate(msk_data_inicial) + 1, "ddmmyy")
    Print #1, x_dados
End Sub
Private Sub FazDetalheAVista(x_empresa As Integer)
    Dim x_dados As String
    With tbl_movimento_cheque_avista
        If .RecordCount > 0 Then
            .Seek ">=", x_empresa, CDate(msk_data_inicial), " ", " ", 0
            If Not .NoMatch Then
                Do Until .EOF
                    If !Empresa <> x_empresa Or ![Data de Emissao] > CDate(msk_data_final) Then
                        Exit Do
                    End If
                    x_dados = Space(48)
                    'Zeros Fixos
                    Mid(x_dados, 1, 1) = "1"
                    'Codigo de Barra 1
                    Mid(x_dados, 2, 8) = ![Codigo de Barra 1]
                    'Codigo de Barra 2
                    Mid(x_dados, 10, 10) = ![Codigo de Barra 2]
                    'Codigo de Barra 3
                    Mid(x_dados, 20, 12) = ![Codigo de Barra 3]
                    'Zeros Fixos
                    Mid(x_dados, 32, 4) = "0000"
                    'Valor do Cheque
                    Mid(x_dados, 36, 13) = Mid(Format(!valor, "00000000000.00"), 1, 11) & Mid(Format(!valor, "00000000000.00"), 13, 2)
                    Print #1, x_dados
                    .MoveNext
                Loop
            End If
        End If
    End With
End Sub
Private Sub Finaliza()
    tbl_movimento_cheque_avista.Close
    
    Set MovCheque = Nothing
End Sub
Function TotalChequeAVista(x_empresa As Integer) As Currency
    Dim i As Integer
    Dim i_inicio As Integer
    Dim i_fim As Integer
    If x_empresa = 0 Then
        i_inicio = 1
        i_fim = 10
    Else
        i_inicio = x_empresa
        i_fim = x_empresa
    End If
    TotalChequeAVista = 0
    For i = i_inicio To i_fim
        With tbl_movimento_cheque_avista
            If .RecordCount > 0 Then
                .Seek ">=", i, CDate(msk_data_inicial), " ", " ", 0
                If Not .NoMatch Then
                    Do Until .EOF
                        If !Empresa <> i Or ![Data de Emissao] > CDate(msk_data_final) Then
                            Exit Do
                        End If
                        TotalChequeAVista = TotalChequeAVista + !valor
                        .MoveNext
                    Loop
                End If
            End If
        End With
    Next
    TotalChequeAVista = TotalChequeAVista + MovCheque.TotalEmissaoPeriodo(x_empresa, CDate(msk_data_inicial.Text), CDate(msk_data_final.Text), "1", "9", "0", "V")
    
End Function
Function ValidaCampos() As Boolean
    ValidaCampos = False
    If Not IsDate(msk_data_inicial) Then
        MsgBox "Informe a data da inicial.", 64, "Atenção!"
        msk_data_inicial.SetFocus
    ElseIf Not IsDate(msk_data_final) Then
        MsgBox "Informe a data da final.", 64, "Atenção!"
        msk_data_final.SetFocus
    ElseIf CDate(msk_data_final) < CDate(msk_data_inicial) Then
        MsgBox "A data final deve ser maior ou igual a " & msk_data_inicial & ".", 64, "Atenção!"
        msk_data_final.SetFocus
    ElseIf Not Val(txt_remessa) > 0 Then
        MsgBox "Informe o número da remessa.", 64, "Atenção!"
        txt_remessa.SetFocus
    ElseIf Not Val(txt_bordero) > 0 Then
        MsgBox "Informe o número do borderô.", 64, "Atenção!"
        txt_bordero.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_ok_Click()
    Dim i As Integer
    Dim x_total As Currency
    If ValidaCampos Then
        CriaDisqueteChequeAVista
        CriaDisqueteChequePreDatado
    End If
End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub Form_Activate()
    msk_data_inicial = Format(CDate(g_data_def) - 1, "dd/mm/yyyy")
    msk_data_final = Format(CDate(g_data_def) - 1, "dd/mm/yyyy")
    txt_remessa = "1"
    txt_bordero = 1
    cmd_ok.SetFocus
    Screen.MousePointer = 1
End Sub
Private Sub Form_Load()
    CentraForm Me
    
    Set tbl_movimento_cheque_avista = bd_sgp.OpenTable("Movimento_Cheque_Avista")
    tbl_movimento_cheque_avista.Index = "id_digitacao"
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_final_GotFocus()
    msk_data_final.SelStart = 0
    msk_data_final.SelLength = 2
End Sub
Private Sub msk_data_final_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_remessa.SetFocus
    End If
End Sub
Private Sub msk_data_inicial_GotFocus()
    msk_data_inicial.SelStart = 0
    msk_data_inicial.SelLength = 2
End Sub
Private Sub msk_data_inicial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data_final.SetFocus
    End If
End Sub
Private Sub txt_bordero_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_bordero_LostFocus()
    txt_bordero = Format(txt_bordero, "000")
End Sub
Private Sub txt_remessa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_bordero.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
