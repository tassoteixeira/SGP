VERSION 5.00
Begin VB.Form gera_string_insert 
   Caption         =   "Gera String do Comando Insert"
   ClientHeight    =   4245
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   6315
   Icon            =   "gera_string_insert.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   6315
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optNuvemNFe 
      Caption         =   "NuvemNFe"
      Height          =   195
      Left            =   3480
      TabIndex        =   24
      Top             =   960
      Width           =   1335
   End
   Begin VB.CheckBox chkVbNet2010 
      Caption         =   "VB.Net 2010"
      Height          =   195
      Left            =   2940
      TabIndex        =   23
      Top             =   2820
      Width           =   2175
   End
   Begin VB.OptionButton optPostgre 
      Caption         =   "Postgre"
      Height          =   195
      Left            =   1800
      TabIndex        =   22
      Top             =   960
      Width           =   1335
   End
   Begin VB.OptionButton optGateData 
      Caption         =   "GateData"
      Height          =   195
      Left            =   60
      TabIndex        =   8
      Top             =   960
      Width           =   1335
   End
   Begin VB.OptionButton optOutroBancoAccess 
      Caption         =   "Outro Banco em Access"
      Height          =   195
      Left            =   60
      TabIndex        =   9
      Top             =   1380
      Width           =   2175
   End
   Begin VB.TextBox txtNomeBancoAccess 
      Height          =   315
      Left            =   2280
      TabIndex        =   11
      Top             =   1860
      Width           =   3975
   End
   Begin VB.CheckBox chkAspNet 
      Caption         =   "ASP.Net"
      Height          =   195
      Left            =   2940
      TabIndex        =   19
      Top             =   3420
      Width           =   2175
   End
   Begin VB.OptionButton optComercial 
      Caption         =   "Comercial"
      Height          =   195
      Left            =   4980
      TabIndex        =   3
      Top             =   240
      Width           =   1335
   End
   Begin VB.OptionButton optTefCerrado 
      Caption         =   "TefCerrado"
      Height          =   195
      Left            =   4980
      TabIndex        =   7
      Top             =   600
      Width           =   1335
   End
   Begin VB.OptionButton optCerradoData 
      Caption         =   "CerradoData"
      Height          =   195
      Left            =   3480
      TabIndex        =   6
      Top             =   600
      Width           =   1335
   End
   Begin VB.OptionButton opt_sfa_data 
      Caption         =   "SFA_DATA"
      Height          =   195
      Left            =   3480
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin VB.OptionButton optSgleData 
      Caption         =   "Sgle_Data"
      Height          =   195
      Left            =   1800
      TabIndex        =   5
      Top             =   600
      Width           =   1335
   End
   Begin VB.OptionButton opt_sgc_data 
      Caption         =   "SGC_DATA"
      Height          =   195
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.OptionButton optDadosInternet 
      Caption         =   "Dados (internet)"
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   600
      Width           =   1575
   End
   Begin VB.CheckBox chkGeraEventos 
      Caption         =   "Gera Eventos VB.Net"
      Height          =   195
      Left            =   1440
      TabIndex        =   20
      Top             =   3720
      Width           =   2175
   End
   Begin VB.CheckBox chkVbNet 
      Caption         =   "VB.Net"
      Height          =   195
      Left            =   2940
      TabIndex        =   16
      Top             =   3120
      Width           =   2175
   End
   Begin VB.OptionButton opt_sgp_data 
      Caption         =   "SGP_DATA"
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   240
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.TextBox txt_quebra_variavel 
      Height          =   315
      Left            =   1440
      MaxLength       =   2
      TabIndex        =   18
      Top             =   3120
      Width           =   375
   End
   Begin VB.TextBox txt_quebra_campo 
      Height          =   315
      Left            =   1440
      MaxLength       =   2
      TabIndex        =   15
      Top             =   2760
      Width           =   375
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&OK"
      Height          =   675
      Left            =   5280
      TabIndex        =   21
      Top             =   3420
      Width           =   915
   End
   Begin VB.TextBox txt_nome_tabela 
      Height          =   315
      Left            =   1440
      TabIndex        =   13
      Top             =   2400
      Width           =   4035
   End
   Begin VB.Label Label4 
      Caption         =   "Nome do &Banco em Access"
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   1860
      Width           =   2175
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   6300
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label Label3 
      Caption         =   "&Quebra Variáveis"
      Height          =   255
      Left            =   60
      TabIndex        =   17
      Top             =   3180
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "&Quebra Campos"
      Height          =   255
      Left            =   60
      TabIndex        =   14
      Top             =   2820
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "&Nome da Tabela:"
      Height          =   255
      Left            =   60
      TabIndex        =   12
      Top             =   2460
      Width           =   1335
   End
End
Attribute VB_Name = "gera_string_insert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lArqTxt As New FileSystemObject
Dim lArquivo As TextStream
Dim lArquivoDestino As TextStream
Dim lNomeProgramaFonte As String
Dim lNomeGroupBox As String
Dim lQtdObjetos As Integer
Dim lRSCriado As Boolean
Dim rsObjeto As New adodb.Recordset

Dim fld As adodb.Field
Dim rst As New adodb.Recordset
Dim lQtdIndice As Integer
Dim lNomeCampoIndice(0 To 10) As String
Dim lNomeCampoIndiceParentese(0 To 10) As String
Dim lNomeCampoIndiceAjustado(0 To 10) As String
Dim lTipoCampoIndice(0 To 10) As String
Dim lNomeTabela As String
Dim lNomeRS As String
Dim l_campo As String
Dim l_arquivo As String
Dim l_condicao As String
Dim l_ordem As String
Dim l_sql As String

Dim cnnSGPDados As New adodb.Connection

Private Sub AtribuiOrdemObjetos()
Dim xString As String
Dim xInicio As Boolean
Dim xVetor As Variant
Dim i As Integer
Dim i2 As Integer

On Error GoTo FileError
    xInicio = False
    i2 = -1
    
    If lArqTxt.FileExists(lNomeProgramaFonte) Then
        Set lArquivo = lArqTxt.OpenTextFile(lNomeProgramaFonte, ForReading)
    
        Do Until lArquivo.AtEndOfStream
            xString = lArquivo.ReadLine
            If xInicio = False Then
                If xString Like "*Windows Form Designer*" Then
                    xInicio = True
                End If
            Else
                If xString = "#End Region" Then
                    Exit Do
                End If
                If xString Like "*.TabIndex = *" Then
                
                
                
                    xVetor = Split(xString, ".")
                    'For i = 0 To lQtdObjetos - 1
                    '    If lNomeObjetos(i) = xVetor(1) Then
                    '        xVetor = Split(xString, " = ")
                    '        lOrdemObjetos(i) = Val(xVetor(1))
                            
                            rsObjeto.Sort = "Nome"
                            rsObjeto.MoveFirst
                            rsObjeto.Find "Nome='" & xVetor(1) & "'"
                            If rsObjeto.EOF = False Then
                                xVetor = Split(xString, " = ")
                                rsObjeto!Ordem = Format(Val(xVetor(1)), "0000")
                                rsObjeto.Update
                            End If
                            
                    '    End If
                    'Next
                End If
            End If
        Loop
        lArquivo.Close
        lArquivoDestino.WriteLine ("  ")
        lArquivoDestino.WriteLine ("Relação dos Objetos Ordenados")
        lArquivoDestino.WriteLine ("  ")
        rsObjeto.Sort = "Ordem"
        rsObjeto.MoveFirst
        Do Until rsObjeto.EOF
            lArquivoDestino.WriteLine (rsObjeto!Nome & "    " & rsObjeto!Ordem)
            rsObjeto.MoveNext
        Loop
    Else
        MsgBox "O programa " & lNomeProgramaFonte & ", não existe!", vbExclamation, "Erro de Verificação"
    End If
    Exit Sub

FileError:
End Sub
Private Sub AtualizaRecordset()
Dim i As Integer
    
    
    'Pega Nome da Tabela
    lNomeTabela = txt_nome_tabela.Text
    lNomeRS = ""
    For i = 1 To Len(txt_nome_tabela.Text)
        If Mid(txt_nome_tabela.Text, i, 9) = "Movimento" Then
            lNomeRS = lNomeRS & Mid(txt_nome_tabela.Text, i, 3)
            i = i + 8
        ElseIf Mid(txt_nome_tabela.Text, i, 1) <> "_" Then
            lNomeRS = lNomeRS & Mid(txt_nome_tabela.Text, i, 1)
        End If
    Next
    
    If optPostgre.Value = True Then
        l_sql = "SELECT * FROM " & lNomeTabela & " LIMIT 1"
    Else
        l_sql = "SELECT TOP 1 * FROM " & lNomeTabela
    End If
    
    On Error GoTo FileError
    rst.CursorLocation = adUseClient
    If opt_sgp_data.Value = True Or optNuvemNFe.Value = True Or opt_sgc_data.Value = True Or opt_sfa_data.Value = True Or optCerradoData.Value = True Or optTefCerrado.Value = True Or optOutroBancoAccess.Value = True Or optGateData.Value = True Or optPostgre.Value = True Then
        Set rst = ConexaoAuxiliar.RsConexao(l_sql)
        'rst.Open l_sql, cnnSGP, adOpenForwardOnly, adLockReadOnly
    ElseIf optDadosInternet.Value = True Or optSgleData.Value = True Or optComercial.Value = True Then
        rst.Open l_sql, cnnSGPDados, adOpenForwardOnly, adLockReadOnly
    End If
    Exit Sub
FileError:
    rst.Close
    rst.CursorLocation = adUseClient
    rst.Open l_sql, cnnSGP, adOpenForwardOnly, adLockReadOnly
    Exit Sub
End Sub
Function CarregaObjetos() As Boolean
Dim xString As String
Dim xInicio As Boolean
Dim xVetor As Variant
Dim i As Integer
Dim i2 As Integer

On Error GoTo FileError
    CarregaObjetos = False
    xInicio = False
    i2 = -1
    lQtdObjetos = 0
    
    'Cria RecordSet
    With rsObjeto
        If lRSCriado Then
            If .RecordCount > 0 Then
                .MoveFirst
                Do Until .EOF
                    .Delete
                    .MoveNext
                Loop
            End If
        Else
            .CursorLocation = adUseClient
            .Fields.Append "Nome", adVarChar, 60
            .Fields.Append "Ordem", adVarChar, 4
            .Open
            lRSCriado = True
        End If
    End With
    
    
    
    
    If lArqTxt.FileExists(lNomeProgramaFonte) Then
        Set lArquivo = lArqTxt.OpenTextFile(lNomeProgramaFonte, ForReading)
    
        Do Until lArquivo.AtEndOfStream
            xString = lArquivo.ReadLine
            If xInicio = False Then
                If xString Like "*Windows Form Designer*" Then
                    xInicio = True
                End If
            Else
                If xString = "#End Region" Then
                    Exit Do
                End If
                If xString Like "*" & lNomeGroupBox & ".Controls.Add*" Then
                    xVetor = Split(xString, ".")
                    For i = LBound(xVetor) To UBound(xVetor)
                        If xVetor(i) = "Add(Me" Then
                            If Mid(xVetor(i + 1), 1, 3) = "txt" Or Mid(xVetor(i + 1), 1, 3) = "cbo" Then
                                i2 = i2 + 1
                                rsObjeto.AddNew
                                rsObjeto!Nome = Mid(xVetor(i + 1), 1, Len(xVetor(i + 1)) - 1)
                                rsObjeto!Ordem = "0000"
                                rsObjeto.Update
                            End If
                        End If
                    Next
                End If
            End If
        Loop
        lArquivo.Close
        If i2 > -1 Then
            lQtdObjetos = i2 + 1
            CarregaObjetos = True
        End If
    Else
        MsgBox "O programa " & lNomeProgramaFonte & ", não existe!", vbExclamation, "Erro de Verificação"
    End If
    Exit Function

FileError:
End Function
Function CriaArquivoTexto() As Boolean
On Error GoTo FileError
    CriaArquivoTexto = False
    Open "\VB5\SGP\DATA\modulo_de_classe.txt" For Output As #1
    CriaArquivoTexto = True
    Exit Function
FileError:
End Function
Function FechaArquivoTexto() As Boolean
Dim retval As Long
On Error GoTo FileError
    FechaArquivoTexto = False
    Close #1
    retval = Shell("c:\WINDOWS\NOTEPAD.EXE \VB5\SGP\DATA\modulo_de_classe.txt", 1)
    FechaArquivoTexto = True
    Exit Function
FileError:
End Function
Private Sub Finaliza()
    'cnnSGP.Close
End Sub
Private Sub GeraArquivoStringInsert()
    Dim retval
    Dim i As Integer
    Dim i2 As Integer
    Dim xString As String
    
    
    i = 1
    For Each fld In rst.Fields
        l_sql = "    x_" & LCase(fld.name) & " = Mid(dados, "
        If fld.Type = 131 Then
            If fld.NumericScale = 0 Then
                i2 = fld.Precision
                l_sql = l_sql & i & ", " & i2 & ")"
            Else
                i2 = fld.Precision - fld.NumericScale
                l_sql = l_sql & i & ", " & i2 & ")"
                i = i + i2
                i2 = fld.NumericScale
                l_sql = l_sql & " & ""."" & MID(DADOS, " & i & ", " & i2 & ")"
            End If
        ElseIf (fld.Type = 200 Or fld.Type = 129) Then
            i2 = fld.DefinedSize
            l_sql = l_sql & i & ", " & i2 & ")"
        Else
            i2 = fld.DefinedSize
            l_sql = l_sql & i & ", " & i2 & ")"
        End If
        Print #1, l_sql
        i = i + i2
    Next
    
    
    
    
    Print #1, l_sql
End Sub
Function GeraDeclaracaoNet(ByVal pVbNet2010 As Boolean) As Boolean
Dim i As Integer
Dim i2 As Integer
Dim xNomeCampo As String

On Error GoTo FileError
    GeraDeclaracaoNet = False
    
    'Print #1, "Option Strict Off"
    'Print #1, "Option Explicit On"
    'Print #1, "Imports System.Data.OleDb"
    Print #1, "Imports System.Data.SqlClient"
    Print #1, ""
    Print #1, "Public Class c" & lNomeRS
    Print #1, ""
    Print #1, "#Region " & Chr(34) & " Declaração " & Chr(34)
    Print #1, ""
    
    If pVbNet2010 = False Then
        For Each fld In rst.Fields
            xNomeCampo = ""
            For i2 = 1 To Len(fld.name)
                If Mid(fld.name, i2, 4) = " da " Then
                    i2 = i2 + 3
                ElseIf Mid(fld.name, i2, 4) = " de " Then
                    i2 = i2 + 3
                ElseIf Mid(fld.name, i2, 4) = " do " Then
                    i2 = i2 + 3
                ElseIf Mid(fld.name, i2, 1) <> "_" And Mid(fld.name, i2, 1) <> " " Then
                    xNomeCampo = xNomeCampo & Mid(fld.name, i2, 1)
                End If
            Next
            l_sql = "    Private m" & xNomeCampo & " As "
            l_sql = l_sql & PreparaTipoCampo(fld)
'            If fld.Type = 131 Then
'                If fld.NumericScale = 2 Then
'                    l_sql = l_sql & "String"
'                Else
'                    If fld.Precision <= 3 Then
'                        l_sql = l_sql & "Short"
'                    Else
'                        l_sql = l_sql & "Integer"
'                    End If
'                End If
'            ElseIf fld.Type = vbInteger Then
'                l_sql = l_sql & "Short"
'            ElseIf fld.Type = vbBoolean Then
'                l_sql = l_sql & "Boolean"
'            ElseIf fld.Type = vbLong Then
'                l_sql = l_sql & "Integer"
'            ElseIf fld.Type = vbCurrency Then
'                l_sql = l_sql & "Decimal"
'            ElseIf fld.Type = vbDate Or fld.Type = 135 Then
'                l_sql = l_sql & "Date"
'            ElseIf (fld.Type = 200 Or fld.Type = 129) Then
'                l_sql = l_sql & "String"
'            Else
'                l_sql = l_sql & "String"
'            End If
            Print #1, l_sql
        Next
        
        Print #1, ""
        If chkAspNet Then
            Print #1, "    Dim oleConn As New OleDbConnection(gStringConexao)"
        Else
            Print #1, "    Dim daTabela As OleDbDataAdapter"
        End If
        Print #1, "    Dim drTabela As OleDbDataReader"
        Print #1, "    Dim cmd As OleDbCommand"
    End If
    Print #1, ""
    Print #1, "#End Region"
    
    GeraDeclaracaoNet = True
    Exit Function
FileError:
End Function
Function GeraDeclaracao() As Boolean
Dim i As Integer
Dim i2 As Integer
Dim xNomeCampo As String

On Error GoTo FileError
    GeraDeclaracao = False
    
    Print #1, "Option Explicit"
    Print #1, ""
    'Print #1, "Private cConexao As adodb.Connection"
    
    For Each fld In rst.Fields
        xNomeCampo = ""
        For i2 = 1 To Len(fld.name)
            If Mid(fld.name, i2, 4) = " da " Then
                i2 = i2 + 3
            ElseIf Mid(fld.name, i2, 4) = " de " Then
                i2 = i2 + 3
            ElseIf Mid(fld.name, i2, 4) = " do " Then
                i2 = i2 + 3
            ElseIf Mid(fld.name, i2, 1) <> "_" And Mid(fld.name, i2, 1) <> " " Then
                xNomeCampo = xNomeCampo & Mid(fld.name, i2, 1)
            End If
        Next
        l_sql = "Private m" & xNomeCampo & " As "
        If fld.Type = 131 Then
            If fld.NumericScale = 2 Then
                l_sql = l_sql & "STRING"
            Else
                If fld.Precision <= 3 Then
                    l_sql = l_sql & "INTEGER"
                Else
                    l_sql = l_sql & "LONG"
                End If
            End If
        ElseIf fld.Type = vbInteger Then
            l_sql = l_sql & "Integer"
        ElseIf fld.Type = vbBoolean Then
            l_sql = l_sql & "Boolean"
        ElseIf fld.Type = vbLong Then
            l_sql = l_sql & "Long"
        ElseIf fld.Type = vbCurrency Then
            l_sql = l_sql & "Currency"
        ElseIf fld.Type = vbDate Or fld.Type = 135 Then
            l_sql = l_sql & "Date"
        ElseIf (fld.Type = 200 Or fld.Type = 129) Then
            l_sql = l_sql & "String"
        Else
            l_sql = l_sql & "String"
        End If
        Print #1, l_sql
    Next
    
    Print #1, ""
    Print #1, "Private rs" & lNomeRS & " As New adodb.Recordset"
    
    GeraDeclaracao = True
    Exit Function
FileError:
End Function
Private Sub GeraEventosObjetos(ByVal pNomeObjeto As String, ByVal pTipoCampo As String)
    'Cria Evento "ENTER"
    lArquivoDestino.WriteLine ("    Private Sub " & pNomeObjeto & "_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles " & pNomeObjeto & ".Enter")
    If pTipoCampo = "Caixa de Texto (String)" Or pTipoCampo = "Caixa de Texto (Inteiro)" Or pTipoCampo = "Caixa de Texto (Valor)" Then
        lArquivoDestino.WriteLine ("        " & pNomeObjeto & ".SelectAll()")
    ElseIf pTipoCampo = "Caixa de Texto (Data)" Then
        lArquivoDestino.WriteLine ("        " & pNomeObjeto & ".Text = fDesmascaraData(") & pNomeObjeto & ".Text)"
        lArquivoDestino.WriteLine ("        " & pNomeObjeto & ".SelectionStart = 0")
        lArquivoDestino.WriteLine ("        " & pNomeObjeto & ".SelectionLength = 4")
    End If
    lArquivoDestino.WriteLine ("    End Sub")
    
    'Cria Evento "KeyPress"
    lArquivoDestino.WriteLine ("    Private Sub " & pNomeObjeto & "_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles " & pNomeObjeto & ".KeyPress")
    lArquivoDestino.WriteLine ("        Dim KeyAscii As Short = Asc(e.KeyChar)")
    lArquivoDestino.WriteLine ("        If KeyAscii = 13 Then")
    lArquivoDestino.WriteLine ("            KeyAscii = 0")
    lArquivoDestino.WriteLine ("            " & ProximoObjeto(rsObjeto!Ordem) & ".Focus()")
    lArquivoDestino.WriteLine ("        End If")
    If pTipoCampo = "Caixa de Texto (Valor)" Then
        lArquivoDestino.WriteLine ("        Call ValidaValor(KeyAscii)")
    ElseIf pTipoCampo = "Caixa de Texto (Inteiro)" Or pTipoCampo = "Caixa de Texto (Data)" Then
        lArquivoDestino.WriteLine ("        Call ValidaInteiro(KeyAscii)")
    End If
    lArquivoDestino.WriteLine ("        If KeyAscii = 0 Then")
    lArquivoDestino.WriteLine ("            e.Handled = True")
    lArquivoDestino.WriteLine ("        End If")
    lArquivoDestino.WriteLine ("    End Sub")
    

    
    'Cria Evento "Leave"
    If pTipoCampo = "Caixa de Texto (Valor)" Or pTipoCampo = "Caixa de Texto (Data)" Then
        lArquivoDestino.WriteLine ("    Private Sub " & pNomeObjeto & "Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles " & pNomeObjeto & ".Leave")
    End If
    If pTipoCampo = "Caixa de Texto (Valor)" Then
        lArquivoDestino.WriteLine ("        If " & pNomeObjeto & ".Text.Trim <> "" Then")
        lArquivoDestino.WriteLine ("            " & pNomeObjeto & ".Text = FormatNumber(" & pNomeObjeto & ".Text, 2)")
        lArquivoDestino.WriteLine ("        End If")
    End If
    If pTipoCampo = "Caixa de Texto (Data)" Then
        lArquivoDestino.WriteLine ("        " & pNomeObjeto & ".Text = fMascaraData(" & pNomeObjeto & ".Text)")
    End If
    If pTipoCampo = "Caixa de Texto (Valor)" Or pTipoCampo = "Caixa de Texto (Data)" Then
        lArquivoDestino.WriteLine ("    End Sub")
    End If
End Sub
Function GeraFuncoesInternas() As Boolean
Dim i As Integer
Dim i2 As Integer
Dim xString As String
Dim xNomeCampo As String
Dim xNomeCampo2 As String

On Error GoTo FileError
    GeraFuncoesInternas = False
    
    Print #1, ""
    Print #1, ""
    Print #1, "'Funções / Procedures internas"
    Print #1, "Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String"
    l_sql = "    PreparaSQL = Trim(" & Chr(34) & "SELECT"
    i = 0
    For Each fld In rst.Fields
        xString = fld.name
        If xString Like "* *" Then
            xString = "[" & xString & "]"
        End If
        If i > 0 Then
            l_sql = l_sql & ","
        End If
        l_sql = l_sql & " " & xString
        i = i + 1
    Next
    l_sql = l_sql & " FROM " & lNomeTabela & Chr(34) & " & " & Chr(34) & " " & Chr(34) & " & xCondicao & " & Chr(34) & " " & Chr(34) & " & xTipoOrdem)"
    Print #1, l_sql
    Print #1, "End Function"
    
    
    
    Print #1, ""
    Print #1, ""
    Print #1, "Private Function AtualizaRecordset(ByVal pQtdRegistro As Integer) As Boolean"
    Print #1, "    Dim i As Integer"
    Print #1, "    AtualizaRecordset = False"
    Print #1, "    Set rs" & lNomeRS & " = New adodb.Recordset"
    Print #1, "    rs" & lNomeRS & ".CursorLocation = adUseClient"
    Print #1, "    i = Len(gSQL)"
    Print #1, "    If pQtdRegistro > 0 Then"
    Print #1, "        gSQL = Mid(gSQL, 1, 6) & " & Chr(34) & " TOP " & Chr(34) & " & pQtdRegistro & Mid(gSQL, 7, i - 6)"
    Print #1, "    End If"
    Print #1, "    rs" & lNomeRS & ".Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly"
    Print #1, "    If Not rs" & lNomeRS & ".EOF Then"
    Print #1, "        AtualizaRecordset = True"
    Print #1, "    End If"
    Print #1, "End Function"
    
    
    
    Print #1, ""
    Print #1, ""
    Print #1, "Private Sub AtribuiValor()"
    For Each fld In rst.Fields
        xNomeCampo = fld.name
        xNomeCampo2 = ""
        For i2 = 1 To Len(fld.name)
            If Mid(fld.name, i2, 4) = " da " Then
                i2 = i2 + 3
            ElseIf Mid(fld.name, i2, 4) = " de " Then
                i2 = i2 + 3
            ElseIf Mid(fld.name, i2, 4) = " do " Then
                i2 = i2 + 3
            ElseIf Mid(fld.name, i2, 1) <> "_" And Mid(fld.name, i2, 1) <> " " Then
                xNomeCampo2 = xNomeCampo2 & Mid(fld.name, i2, 1)
            End If
        Next
        l_sql = "    " & xNomeCampo2 & " = rs" & lNomeRS & "(" & Chr(34) & xNomeCampo & Chr(34) & ").Value"
        Print #1, l_sql
    Next
    Print #1, "End Sub"
    
    
    
    Print #1, ""
    Print #1, ""
    Print #1, "Private Function Localizar(ByVal pQtdRegistro As Integer) As Boolean"
    Print #1, "    Localizar = False"
    Print #1, "    If AtualizaRecordset(pQtdRegistro) Then"
    Print #1, "        Localizar = True"
'    Print #1, "        AtribuiValor"
    Print #1, "    End If"
'    Print #1, "    rs" & lNomeRS & ".Close"
'    Print #1, "    Set rs" & lNomeRS & " = Nothing"
    Print #1, "End Function"
    
    
    
    GeraFuncoesInternas = True
    Exit Function
FileError:
End Function
Function GeraFuncoesInternasNet(ByVal pVbNet2010 As Boolean) As Boolean
Dim i As Integer
Dim i2 As Integer
Dim xString As String
Dim xNomeCampo As String
Dim xNomeCampo2 As String

On Error GoTo FileError
    GeraFuncoesInternasNet = False
    
    Print #1, "    Public Function MontaDS(ByVal pSQL As String) As DataSet"
    Print #1, "        Dim dsTabela As New DataSet"
    If pVbNet2010 Then
        Print #1, "        Dim daTabela As New SqlDataAdapter"
        Print #1, ""
        Print #1, "        Try"
        Print #1, "            daTabela = New SqlDataAdapter(pSQL, gBdFuncoesDiversas.gConnAzure)"
        Print #1, "            daTabela.Fill(dsTabela, " & Chr(34) & lNomeTabela & Chr(34) & ")"
        Print #1, "            Return dsTabela"
        Print #1, "        Catch"
        Print #1, "            gFuncoesDiversas.CriaLog(Me.GetType.Name & "":MontaDS - Erro não identificado."", Err.Description, pSQL)"
        Print #1, "            Return dsTabela"
    Else
        Print #1, ""
        Print #1, "        Try"
        Print #1, "            daTabela = New OleDbDataAdapter(pSQL, gConn)"
        Print #1, "            daTabela.Fill(dsTabela, " & Chr(34) & lNomeTabela & Chr(34) & ")"
        Print #1, "            Return dsTabela"
        Print #1, "        Catch"
        Print #1, "            CriaLogRN(Me.GetType.Name & "":MontaDS - Erro não identificado."", Err.Description, pSQL)"
        Print #1, "        Finally"
    End If
    Print #1, "        End Try"
    Print #1, "    End Function"
    Print #1, ""
    Print #1, "#End Region"
    Print #1, ""
    Print #1, "#Region " & Chr(34) & " Funções/Procedures Internas da Classe " & Chr(34)
    Print #1, ""
    
    
    Print #1, "    Private Sub AtribuiValor()"
    Print #1, "        Dim xLocal As Short"
    Print #1, "        Try"
    i = 0
    For Each fld In rst.Fields
        i = i + 1
        xNomeCampo = fld.name
        xNomeCampo2 = PreparaNomePropriedade(fld.name)
'        For i2 = 1 To Len(fld.name)
'            If Mid(fld.name, i2, 4) = " da " Then
'                i2 = i2 + 3
'            ElseIf Mid(fld.name, i2, 4) = " de " Then
'                i2 = i2 + 3
'            ElseIf Mid(fld.name, i2, 4) = " do " Then
'                i2 = i2 + 3
'            ElseIf Mid(fld.name, i2, 1) <> "_" And Mid(fld.name, i2, 1) <> " " Then
'                xNomeCampo2 = xNomeCampo2 & Mid(fld.name, i2, 1)
'            End If
'        Next
        If pVbNet2010 Then
            l_sql = "            " & xNomeCampo2 & " = gBdFuncoesDiversas.gDrTabelaAzure.Item(" & Chr(34) & xNomeCampo & Chr(34) & ")"
        Else
            l_sql = "            " & xNomeCampo2 & " = drTabela.Item(" & Chr(34) & xNomeCampo & Chr(34) & ")"
        End If
        Print #1, "            xLocal = " & i
        Print #1, l_sql
    Next
    Print #1, "        Catch ex As Exception"
    If pVbNet2010 Then
        Print #1, "            gFuncoesDiversas.CriaLog(Me.GetType.Name & "":AtribuiValor - Erro não identificado."", Err.Description, ""xLocal="" & xLocal)"
    Else
        Print #1, "            CriaLogRN(Me.GetType.Name & "":AtribuiValor - Erro não identificado."", Err.Description, ""xLocal="" & xLocal)"
    End If
    Print #1, "        End Try"
    Print #1, "    End Sub"
    
    
    Print #1, "    Private Function Localizar(ByVal pQtdRegistro As Short, ByVal pLeRegistro As Boolean, ByVal pAtribuiValor As Boolean, ByVal pFecha As Boolean) As Boolean"
    If pVbNet2010 Then
        Print #1, "        Localizar = False"
        Print #1, "        If gBdFuncoesDiversas.LocalizarRegistroAzure(pQtdRegistro, pLeRegistro, pAtribuiValor, pFecha, Me.GetType.Name & "":Localizar:"") Then"
        Print #1, "            Localizar = True"
        Print #1, "            If pAtribuiValor Then"
        Print #1, "                AtribuiValor()"
        Print #1, "                If pFecha Then"
        Print #1, "                    gBdFuncoesDiversas.FechaCmdDataReaderAzure(True, True)"
        Print #1, "                End If"
        Print #1, "            End If"
        Print #1, "        End If"
    Else
        Print #1, "        Dim i As Short"
        Print #1, ""
        Print #1, "        Localizar = False"
        Print #1, "        i = Len(gSQL)"
        Print #1, "        If pQtdRegistro > 0 Then"
        Print #1, "            gSQL = Mid(gSQL, 1, 6) & "" Top "" & pQtdRegistro & Mid(gSQL, 7, i - 6)"
        Print #1, "        End If"
        Print #1, "        Try"
        Print #1, "            cmd = New OleDbCommand(gSQL, gConn)"
        Print #1, "            drTabela = cmd.ExecuteReader"
        Print #1, "            If pLeRegistro = True Then"
        Print #1, "                If drTabela.Read() Then"
        Print #1, "                    Localizar = True"
        Print #1, "                    If pAtribuiValor Then"
        Print #1, "                        AtribuiValor()"
        Print #1, "                    End If"
        Print #1, "                End If"
        Print #1, "            Else"
        Print #1, "                Localizar = True"
        Print #1, "            End If"
        Print #1, "        Catch"
        Print #1, "            CriaLogRN(Me.GetType.Name & "":Localizar - Erro não identificado."", Err.Description, gSQL)"
        Print #1, "        Finally"
        Print #1, "            If pFecha Then"
        Print #1, "                drTabela.Close()"
        Print #1, "                cmd.Dispose()"
        Print #1, "            End If"
        Print #1, "        End Try"
    End If
    Print #1, "    End Function"
    
    
    If pVbNet2010 Then
        Print #1, "    Private Sub PreparaSbSQL(ByVal pCondicao As String, ByVal pTipoOrdem As String)"
        Print #1, "        sbSQL.Clear()"
        l_sql = "        sbSQL.Append(" & Chr(34) & "SELECT"
        i = 0
        For Each fld In rst.Fields
            xString = fld.name
            If xString Like "* *" Then
                xString = "[" & xString & "]"
            End If
            If i > 0 Then
                l_sql = l_sql & ","
            End If
            l_sql = l_sql & " " & xString
            i = i + 1
        Next
        Print #1, l_sql & Chr(34) & ")"
        Print #1, "        sbSQL.Append(" & Chr(34) & " FROM " & lNomeTabela & Chr(34) & ")"
        Print #1, "        sbSQL.Append(" & Chr(34) & " " & Chr(34) & ")"
        Print #1, "        sbSQL.Append(pCondicao)"
        Print #1, "        sbSQL.Append(" & Chr(34) & " " & Chr(34) & ")"
        Print #1, "        sbSQL.Append(pTipoOrdem)"
        Print #1, "    End Function"
    Else
        Print #1, "    Private Function PreparaSQL(ByVal pCondicao As String, ByVal pTipoOrdem As String) As String"
        l_sql = "        PreparaSQL = Trim(" & Chr(34) & "SELECT"
        i = 0
        For Each fld In rst.Fields
            xString = fld.name
            If xString Like "* *" Then
                xString = "[" & xString & "]"
            End If
            If i > 0 Then
                l_sql = l_sql & ","
            End If
            l_sql = l_sql & " " & xString
            i = i + 1
        Next
        l_sql = l_sql & " FROM " & lNomeTabela & Chr(34) & " & " & Chr(34) & " " & Chr(34) & " & pCondicao & " & Chr(34) & " " & Chr(34) & " & pTipoOrdem)"
        Print #1, l_sql
        Print #1, "    End Function"
    End If
    
    
    Print #1, ""
    Print #1, "#End Region"
    Print #1, ""
    Print #1, "End Class"
    
    
    
    GeraFuncoesInternasNet = True
    Exit Function
FileError:
End Function
Function GeraMetodoAlterar() As Boolean
Dim i As Integer
Dim i2 As Integer
Dim xString As String
Dim xVirgula As String
Dim xNomeCampo As String
Dim xNomeCampo2 As String

On Error GoTo FileError
    GeraMetodoAlterar = False
    Print #1, ""
    Print #1, ""
    xString = "Public Function Alterar("
    For i = 0 To lQtdIndice
        If i > 0 Then
            xString = xString & ", "
        End If
        xString = xString & "ByVal p" & lNomeCampoIndiceAjustado(i) & " As " & lTipoCampoIndice(i)
    Next
    xString = xString & ") As Boolean"
    Print #1, xString
    Print #1, "Dim lRecordsAffected As Long"
    Print #1, ""
    Print #1, "On Error GoTo Err_Alterar"
    Print #1, ""
    Print #1, "    Alterar = False"
    l_sql = "    gSQL = " & Chr(34) & "UPDATE " & lNomeTabela & " SET " & Chr(34)
    Print #1, l_sql
    i = 0
    For Each fld In rst.Fields
        i = i + 1
        xNomeCampo = fld.name
        If xNomeCampo Like "* *" Then
            xNomeCampo = "[" & xNomeCampo & "]"
        End If
        xNomeCampo2 = ""
        For i2 = 1 To Len(fld.name)
            If Mid(fld.name, i2, 4) = " da " Then
                i2 = i2 + 3
            ElseIf Mid(fld.name, i2, 4) = " de " Then
                i2 = i2 + 3
            ElseIf Mid(fld.name, i2, 4) = " do " Then
                i2 = i2 + 3
            ElseIf Mid(fld.name, i2, 1) <> "_" And Mid(fld.name, i2, 1) <> " " Then
                xNomeCampo2 = xNomeCampo2 & Mid(fld.name, i2, 1)
            End If
        Next
        If bdOracle Then
        ElseIf bdAccess Then
            'ACCESS INTEGER
            If fld.Type = vbInteger Then
                xString = "sqlNumero"
            ElseIf fld.Type = vbLong Then
                xString = "sqlNumero"
            ElseIf fld.Type = vbCurrency Then
                xString = "sqlValor"
            ElseIf fld.Type = 135 Then
                xString = "sqlData"
            ElseIf fld.Type = 129 Or fld.Type = 200 Then
                xString = "sqlTexto"
            ElseIf fld.Type = vbBoolean Then
                xString = "sqlBoolean"
            Else
                xString = "sqlTexto"
            End If
        ElseIf bdSqlServer Then
            If fld.Type = vbInteger Then
                xString = "sqlNumero"
            ElseIf fld.Type = vbLong Then
                xString = "sqlNumero"
            ElseIf fld.Type = vbCurrency Then
                xString = "sqlValor"
            ElseIf fld.Type = 135 Then
                xString = "sqlData"
            ElseIf fld.Type = 129 Or fld.Type = 200 Then
                xString = "sqlTexto"
            ElseIf fld.Type = vbBoolean Then
                xString = "sqlBoolean"
            Else
                xString = "sqlTexto"
            End If
        End If
        xVirgula = ""
        If i > 1 Then
            xVirgula = ", "
        End If
        l_sql = "    Call " & xString & "(2, " & Chr(34) & xVirgula & xNomeCampo & " = " & Chr(34) & ", " & "m" & xNomeCampo2 & ")"
        Print #1, l_sql
    Next
    
    i = 0
    For Each fld In rst.Fields
        If fld.name = lNomeCampoIndice(i) Then
            xNomeCampo = fld.name
            If xNomeCampo Like "* *" Then
                xNomeCampo = "[" & xNomeCampo & "]"
            End If
            xNomeCampo2 = ""
            For i2 = 1 To Len(fld.name)
                If Mid(fld.name, i2, 4) = " da " Then
                    i2 = i2 + 3
                ElseIf Mid(fld.name, i2, 4) = " de " Then
                    i2 = i2 + 3
                ElseIf Mid(fld.name, i2, 4) = " do " Then
                    i2 = i2 + 3
                ElseIf Mid(fld.name, i2, 1) <> "_" And Mid(fld.name, i2, 1) <> " " Then
                    xNomeCampo2 = xNomeCampo2 & Mid(fld.name, i2, 1)
                End If
            Next
            i = i + 1
            If i = 1 Then
                l_sql = "    gSQL = gSQL & " & Chr(34) & " WHERE "
            Else
                l_sql = "    gSQL = gSQL & " & Chr(34) & " AND "
            End If
            If bdOracle Then
            ElseIf bdAccess Then
                'ACCESS INTEGER
                If fld.Type = vbInteger Then
                    l_sql = l_sql & xNomeCampo & " = " & Chr(34) & " & " & "p" & xNomeCampo2
                ElseIf fld.Type = vbLong Then
                    l_sql = l_sql & xNomeCampo & " = " & Chr(34) & " & " & "p" & xNomeCampo2
                ElseIf fld.Type = vbCurrency Then
                    l_sql = l_sql & xNomeCampo & " = " & Chr(34) & " & " & "p" & xNomeCampo2
                ElseIf fld.Type = 135 Or fld.Type = vbDate Then
                    l_sql = l_sql & xNomeCampo & " = " & Chr(34) & " & " & "preparaData(p" & xNomeCampo2 & ")"
                ElseIf fld.Type = 129 Or fld.Type = 200 Then
                    l_sql = l_sql & xNomeCampo & " = " & Chr(34) & " & " & "preparaTexto(p" & xNomeCampo2 & ")"
                Else
                    l_sql = l_sql & xNomeCampo & " = " & Chr(34) & " & " & "p" & xNomeCampo2
                End If
            End If
            Print #1, l_sql
        End If
    Next
    l_sql = "    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords"
    Print #1, l_sql
    Print #1, "    If lRecordsAffected > 0 Then"
    Print #1, "        Alterar = True"
    Print #1, "    End If"
    Print #1, "    Exit Function"
    Print #1, ""
    Print #1, "Err_Alterar:"
    Print #1, "End Function"
    GeraMetodoAlterar = True
    Exit Function
FileError:
End Function
Function GeraMetodoAlterarNet(ByVal pVbNet2010 As Boolean) As Boolean
Dim i As Integer
Dim i2 As Integer
Dim xString As String
Dim xVirgula As String
Dim xNomeCampo As String
Dim xNomeCampo2 As String
Dim xTipoCampo As String

On Error GoTo FileError
    
    GeraMetodoAlterarNet = False
    Print #1, ""
    Print #1, "#Region " & Chr(34) & " Métodos da Classe " & Chr(34)
    Print #1, ""
    xString = "    Public Function Alterar("
    For i = 0 To lQtdIndice
        If i > 0 Then
            xString = xString & ", "
        End If
        xString = xString & "ByVal p" & lNomeCampoIndiceAjustado(i) & " As " & lTipoCampoIndice(i)
    Next
    xString = xString & ") As Boolean"
    Print #1, xString
    Print #1, "        Alterar = False"
    Print #1, "        Try"
    If chkAspNet Then
        l_sql = "            sbSQL.Remove(0, sbSQL.Length)"
        Print #1, l_sql
        l_sql = "            sbSQL.Append(" & Chr(34) & "UPDATE " & lNomeTabela & " SET " & Chr(34) & ")"
    Else
        'l_sql = "            gSQL = " & Chr(34) & "UPDATE " & lNomeTabela & " SET " & Chr(34)
        l_sql = "            sbSQL.Clear()"
        Print #1, l_sql
        l_sql = "            sbSQL.Append(" & Chr(34) & "UPDATE " & lNomeTabela & " SET " & Chr(34) & ")"
    End If
    Print #1, l_sql
    i = 0
    For Each fld In rst.Fields
        i = i + 1
        xNomeCampo = fld.name
        If xNomeCampo Like "* *" Then
            xNomeCampo = "[" & xNomeCampo & "]"
        End If
        xNomeCampo2 = PreparaNomePropriedade(fld.name)
        xString = PreparaTipoCampo(fld)
        xTipoCampo = PreparaTipoCampo2(fld)
        xVirgula = ""
        If i > 1 Then
            xVirgula = ", "
        End If
        If pVbNet2010 = True Then
            If xTipoCampo = "sqlTexto" Then
                l_sql = "            gBdFuncoesDiversas." & xTipoCampo & "Sb(2, " & Chr(34) & xVirgula & xNomeCampo & " = " & Chr(34) & ", " & xNomeCampo2 & ", gBdFuncoesDiversas.bdEnumAzure)"
            ElseIf xTipoCampo = "sqlBoolean" Then
                l_sql = "            gBdFuncoesDiversas.sqlBoleanoSb(2, " & Chr(34) & xVirgula & xNomeCampo & " = " & Chr(34) & ", " & xNomeCampo2 & ".ToString, gBdFuncoesDiversas.bdEnumAzure)"
            Else
                'l_sql = "            " & xString & "(2, " & Chr(34) & xVirgula & xNomeCampo & " = " & Chr(34) & ", " & xNomeCampo2 & ".ToString)"
                l_sql = "            gBdFuncoesDiversas." & xTipoCampo & "Sb(2, " & Chr(34) & xVirgula & xNomeCampo & " = " & Chr(34) & ", " & xNomeCampo2 & ".ToString, gBdFuncoesDiversas.bdEnumAzure)"
            End If
        Else
            If xString = "sqlTexto" Then
                l_sql = "            " & xString & "(2, " & Chr(34) & xVirgula & xNomeCampo & " = " & Chr(34) & ", " & "m" & xNomeCampo2 & ")"
            Else
                l_sql = "            " & xString & "(2, " & Chr(34) & xVirgula & xNomeCampo & " = " & Chr(34) & ", " & "m" & xNomeCampo2 & ".ToString)"
            End If
        End If
        Print #1, l_sql
    Next
    
    i = 0
    For Each fld In rst.Fields
        If fld.name = lNomeCampoIndice(i) Then
            xNomeCampo = fld.name
            If xNomeCampo Like "* *" Then
                xNomeCampo = "[" & xNomeCampo & "]"
            End If
            xNomeCampo2 = ""
            For i2 = 1 To Len(fld.name)
                If Mid(fld.name, i2, 4) = " da " Then
                    i2 = i2 + 3
                ElseIf Mid(fld.name, i2, 4) = " de " Then
                    i2 = i2 + 3
                ElseIf Mid(fld.name, i2, 4) = " do " Then
                    i2 = i2 + 3
                ElseIf Mid(fld.name, i2, 1) <> "_" And Mid(fld.name, i2, 1) <> " " Then
                    xNomeCampo2 = xNomeCampo2 & Mid(fld.name, i2, 1)
                End If
            Next
            i = i + 1
            If i = 1 Then
                l_sql = "            gSQL += " & Chr(34) & " WHERE "
            Else
                l_sql = "            gSQL += " & Chr(34) & " AND "
            End If
            If bdOracle Then
            ElseIf bdAccess Then
                'ACCESS INTEGER
                If fld.Type = vbInteger Then
                    l_sql = l_sql & xNomeCampo & " = " & Chr(34) & " & " & "p" & xNomeCampo2
                ElseIf fld.Type = vbLong Then
                    l_sql = l_sql & xNomeCampo & " = " & Chr(34) & " & " & "p" & xNomeCampo2
                ElseIf fld.Type = vbCurrency Then
                    l_sql = l_sql & xNomeCampo & " = " & Chr(34) & " & " & "p" & xNomeCampo2
                ElseIf fld.Type = 135 Or fld.Type = vbDate Then
                    l_sql = l_sql & xNomeCampo & " = " & Chr(34) & " & " & "preparaData(p" & xNomeCampo2 & ")"
                ElseIf fld.Type = 129 Or fld.Type = 200 Then
                    l_sql = l_sql & xNomeCampo & " = " & Chr(34) & " & " & "preparaTexto(p" & xNomeCampo2 & ")"
                Else
                    l_sql = l_sql & xNomeCampo & " = " & Chr(34) & " & " & "p" & xNomeCampo2
                End If
            End If
            Print #1, l_sql
        End If
    Next
    If chkAspNet Then
        Print #1, "            oleConn.Open()"
        Print #1, "            cmd = New OleDbCommand(sbSQL.ToString, oleConn)"
    Else
        'Print #1, "            cmd = New OleDbCommand(gSQL, gConn)"
    End If
    'Print #1, "            If cmd.ExecuteNonQuery() > 0 Then"
    'Print #1, "                Alterar = True"
    'Print #1, "            Else"
    Print #1, "            Alterar = gBdFuncoesDiversas.ExecutaCmdAzure(sbSQL.ToString, Me.GetType.Name & " & Chr(34) & ":Alterar" & Chr(34) & ")"
    If chkAspNet Then
        Print #1, "                CriaLogRN(Me.GetType.Name & "":Alterar - Erro ao alterar registro."", ""Err.Description"", sbSQL.ToString)"
    Else
        'Print #1, "                CriaLogRN(Me.GetType.Name & "":Alterar - Erro ao alterar registro."", ""Err.Description"", gSQL)"
    End If
    'Print #1, "            End If"
    Print #1, "        Catch"
    If chkAspNet Then
        Print #1, "            CriaLogRN(Me.GetType.Name & "":Alterar - Erro não identificado."", Err.Description, sbSQL.ToString)"
    Else
        'Print #1, "            CriaLogRN(Me.GetType.Name & "":Alterar - Erro não identificado."", Err.Description, gSQL)"
        Print #1, "            gFuncoesDiversas.CriaLog(Me.GetType.Name & "":Alterar - Erro não identificado."", Err.Description, sbSQL.ToString)"
    End If
    'Print #1, "        Finally"
    'Print #1, "            cmd.Dispose()"
    If chkAspNet Then
        Print #1, "            oleConn.Close()"
    End If
    Print #1, "        End Try"
    Print #1, "    End Function"
    GeraMetodoAlterarNet = True
    Exit Function
FileError:
End Function
Function GeraMetodoExcluir() As Boolean
Dim i As Integer
Dim i2 As Integer
Dim xString As String
Dim xNomeCampo As String
Dim xNomeCampo2 As String

On Error GoTo FileError
    GeraMetodoExcluir = False
    Print #1, ""
    Print #1, ""
    xString = "Public Function Excluir("
    For i = 0 To lQtdIndice
        If i > 0 Then
            xString = xString & ", "
        End If
        xString = xString & "ByVal p" & lNomeCampoIndiceAjustado(i) & " As " & lTipoCampoIndice(i)
    Next
    xString = xString & ") As Boolean"
    Print #1, xString
    Print #1, "Dim lRecordsAffected As Long"
    Print #1, ""
    Print #1, "On Error GoTo Err_Excluir"
    Print #1, ""
    Print #1, "    Excluir = False"
    l_sql = "    gSQL = " & Chr(34) & "DELETE FROM " & lNomeTabela & Chr(34)
    Print #1, l_sql
    i = 0
    For Each fld In rst.Fields
        If fld.name = lNomeCampoIndice(i) Then
            xNomeCampo = fld.name
            If xNomeCampo Like "* *" Then
                xNomeCampo = "[" & xNomeCampo & "]"
            End If
            xNomeCampo2 = ""
            For i2 = 1 To Len(fld.name)
                If Mid(fld.name, i2, 4) = " da " Then
                    i2 = i2 + 3
                ElseIf Mid(fld.name, i2, 4) = " de " Then
                    i2 = i2 + 3
                ElseIf Mid(fld.name, i2, 4) = " do " Then
                    i2 = i2 + 3
                ElseIf Mid(fld.name, i2, 1) <> "_" And Mid(fld.name, i2, 1) <> " " Then
                    xNomeCampo2 = xNomeCampo2 & Mid(fld.name, i2, 1)
                End If
            Next
            i = i + 1
            If i = 1 Then
                l_sql = "    gSQL = gSQL & " & Chr(34) & " WHERE "
            Else
                l_sql = "    gSQL = gSQL & " & Chr(34) & " AND "
            End If
            If bdOracle Then
            ElseIf bdAccess Then
                'ACCESS INTEGER
                If fld.Type = vbInteger Then
                    l_sql = l_sql & xNomeCampo & " = " & Chr(34) & " & " & "p" & xNomeCampo2
                ElseIf fld.Type = vbLong Then
                    l_sql = l_sql & xNomeCampo & " = " & Chr(34) & " & " & "p" & xNomeCampo2
                ElseIf fld.Type = vbCurrency Then
                    l_sql = l_sql & xNomeCampo & " = " & Chr(34) & " & " & "p" & xNomeCampo2
                ElseIf fld.Type = 135 Or fld.Type = vbDate Then
                    l_sql = l_sql & xNomeCampo & " = " & Chr(34) & " & " & "preparaData(p" & xNomeCampo2 & ")"
                ElseIf fld.Type = 129 Or fld.Type = 200 Then
                    l_sql = l_sql & xNomeCampo & " = " & Chr(34) & " & " & "preparaTexto(p" & xNomeCampo2 & ")"
                Else
                    l_sql = l_sql & xNomeCampo & " = " & Chr(34) & " & " & "p" & xNomeCampo2
                End If
            End If
            Print #1, l_sql
        End If
    Next
    l_sql = "    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords"
    Print #1, l_sql
    Print #1, "    If lRecordsAffected > 0 Then"
    Print #1, "        Excluir = True"
    Print #1, "    End If"
    Print #1, "    Exit Function"
    Print #1, ""
    Print #1, "Err_Excluir:"
    Print #1, "End Function"
    GeraMetodoExcluir = True
    Exit Function
FileError:
End Function
Function GeraMetodoExcluirNet() As Boolean
Dim i As Integer
Dim i2 As Integer
Dim xString As String
Dim xNomeCampo As String
Dim xNomeCampo2 As String

On Error GoTo FileError
    GeraMetodoExcluirNet = False
    xString = "    Public Function Excluir("
    For i = 0 To lQtdIndice
        If i > 0 Then
            xString = xString & ", "
        End If
        xString = xString & "ByVal p" & lNomeCampoIndiceAjustado(i) & " As " & lTipoCampoIndice(i)
    Next
    xString = xString & ") As Boolean"
    Print #1, xString
    Print #1, "        Excluir = False"
    Print #1, "        Try"
    If chkAspNet Then
        l_sql = "            sbSQL.Remove(0, sbSQL.Length)"
        Print #1, l_sql
        l_sql = "            sbSQL.Append(" & Chr(34) & "DELETE FROM " & lNomeTabela & Chr(34) & ")"
    Else
        l_sql = "            sbSQL.Clear()"
        Print #1, l_sql
        l_sql = "            sbSQL.Append(" & Chr(34) & "DELETE FROM " & lNomeTabela & Chr(34) & ")"
    End If
    Print #1, l_sql
    i = 0
    For Each fld In rst.Fields
        If fld.name = lNomeCampoIndice(i) Then
            xNomeCampo = fld.name
            If xNomeCampo Like "* *" Then
                xNomeCampo = "[" & xNomeCampo & "]"
            End If
            xNomeCampo2 = ""
            For i2 = 1 To Len(fld.name)
                If Mid(fld.name, i2, 4) = " da " Then
                    i2 = i2 + 3
                ElseIf Mid(fld.name, i2, 4) = " de " Then
                    i2 = i2 + 3
                ElseIf Mid(fld.name, i2, 4) = " do " Then
                    i2 = i2 + 3
                ElseIf Mid(fld.name, i2, 1) <> "_" And Mid(fld.name, i2, 1) <> " " Then
                    xNomeCampo2 = xNomeCampo2 & Mid(fld.name, i2, 1)
                End If
            Next
            i = i + 1
            If i = 1 Then
                l_sql = "            gSQL += " & Chr(34) & " WHERE "
            Else
                l_sql = "            gSQL += " & Chr(34) & " AND "
            End If
            If bdOracle Then
            ElseIf bdAccess Then
                If fld.Type = vbInteger Then
                    l_sql = l_sql & xNomeCampo & " = " & Chr(34) & " & " & "p" & xNomeCampo2
                ElseIf fld.Type = vbLong Then
                    l_sql = l_sql & xNomeCampo & " = " & Chr(34) & " & " & "p" & xNomeCampo2
                ElseIf fld.Type = vbCurrency Then
                    l_sql = l_sql & xNomeCampo & " = " & Chr(34) & " & " & "p" & xNomeCampo2
                ElseIf fld.Type = 135 Or fld.Type = vbDate Then
                    l_sql = l_sql & xNomeCampo & " = " & Chr(34) & " & " & "preparaData(p" & xNomeCampo2 & ")"
                ElseIf fld.Type = 129 Or fld.Type = 200 Then
                    l_sql = l_sql & xNomeCampo & " = " & Chr(34) & " & " & "preparaTexto(p" & xNomeCampo2 & ")"
                Else
                    l_sql = l_sql & xNomeCampo & " = " & Chr(34) & " & " & "p" & xNomeCampo2
                End If
            ElseIf bdSqlServer Then
                If fld.Type = vbInteger Then
                    l_sql = l_sql & xNomeCampo & " = " & Chr(34) & " & " & "p" & xNomeCampo2
                ElseIf fld.Type = vbLong Then
                    l_sql = l_sql & xNomeCampo & " = " & Chr(34) & " & " & "p" & xNomeCampo2
                ElseIf fld.Type = vbCurrency Then
                    l_sql = l_sql & xNomeCampo & " = " & Chr(34) & " & " & "p" & xNomeCampo2
                ElseIf fld.Type = 135 Or fld.Type = vbDate Then
                    l_sql = l_sql & xNomeCampo & " = " & Chr(34) & " & " & "preparaData(p" & xNomeCampo2 & ")"
                ElseIf fld.Type = 129 Or fld.Type = 200 Then
                    l_sql = l_sql & xNomeCampo & " = " & Chr(34) & " & " & "preparaTexto(p" & xNomeCampo2 & ")"
                Else
                    l_sql = l_sql & xNomeCampo & " = " & Chr(34) & " & " & "p" & xNomeCampo2
                End If
            End If
            Print #1, l_sql
        End If
    Next
    If chkAspNet Then
        Print #1, "            oleConn.Open()"
        Print #1, "            cmd = New OleDbCommand(sbSQL.ToString, oleConn)"
    Else
        'Print #1, "            cmd = New OleDbCommand(gSQL, gConn)"
    End If
    'Print #1, "            If cmd.ExecuteNonQuery() > 0 Then"
    'Print #1, "                Excluir = True"
    'Print #1, "            Else"
    Print #1, "            Excluir = gBdFuncoesDiversas.ExecutaCmdAzure(sbSQL.ToString, Me.GetType.Name & " & Chr(34) & ":Excluir" & Chr(34) & ")"
    If chkAspNet Then
        Print #1, "                CriaLogRN(Me.GetType.Name & "":Excluir - Erro ao excluir registro."", ""Err.Description"", sbSQL.ToString)"
    Else
        'Print #1, "                CriaLogRN(Me.GetType.Name & "":Excluir - Erro ao excluir registro."", ""Err.Description"", gSQL)"
    End If
    'Print #1, "            End If"
    Print #1, "        Catch"
    If chkAspNet Then
        Print #1, "            CriaLogRN(Me.GetType.Name & "":Excluir - Erro não identificado."", Err.Description, sbSQL.ToString)"
    Else
        'Print #1, "            CriaLogRN(Me.GetType.Name & "":Excluir - Erro não identificado."", Err.Description, gSQL)"
        Print #1, "            gFuncoesDiversas.CriaLog(Me.GetType.Name & "":Excluir - Erro não identificado."", Err.Description, sbSQL.ToString)"
    End If
    'Print #1, "        Finally"
    'Print #1, "            cmd.Dispose()"
    If chkAspNet Then
        Print #1, "            oleConn.Close()"
    End If
    Print #1, "        End Try"
    Print #1, "    End Function"
    GeraMetodoExcluirNet = True
    Exit Function
FileError:
End Function
Function GeraMetodoIncluir() As Boolean
Dim i As Integer
Dim i2 As Integer
Dim xString As String
Dim xTipoCampo As String

On Error GoTo FileError
    GeraMetodoIncluir = False
    Print #1, ""
    Print #1, ""
    Print #1, "Public Function Incluir() As Boolean"
    Print #1, "Dim lRecordsAffected As Long"
    Print #1, ""
    Print #1, "On Error GoTo Err_Incluir"
    Print #1, ""
    Print #1, "    Incluir = False"
    l_sql = "    gSQL = " & Chr(34) & "INSERT INTO " & lNomeTabela & " ( "
    i = 0
    For Each fld In rst.Fields
        If i = Val(txt_quebra_campo.Text) Then
            l_sql = l_sql & Chr(34)
            Print #1, l_sql
            l_sql = "    gSQL = gSQL & " & Chr(34)
            i = 0
        End If
        xString = fld.name
        If xString Like "* *" Then
            xString = "[" & xString & "]"
        End If
        l_sql = l_sql & xString & ", "
        i = i + 1
    Next
    If i > 0 Then
        l_sql = Mid(l_sql, 1, (Len(l_sql) - 2)) & " ) VALUES ( " & Chr(34)
        Print #1, l_sql
    End If
    l_sql = ""
    i = 0
    For Each fld In rst.Fields
        If i = Val(txt_quebra_variavel.Text) Then
            Print #1, l_sql
            l_sql = "        gSQL = gSQL & "
            i = 0
        End If
        'xString = fld.Name
        xString = ""
        For i2 = 1 To Len(fld.name)
            If Mid(fld.name, i2, 4) = " da " Then
                i2 = i2 + 3
            ElseIf Mid(fld.name, i2, 4) = " de " Then
                i2 = i2 + 3
            ElseIf Mid(fld.name, i2, 4) = " do " Then
                i2 = i2 + 3
            ElseIf Mid(fld.name, i2, 1) <> "_" And Mid(fld.name, i2, 1) <> " " Then
                xString = xString & Mid(fld.name, i2, 1)
            End If
        Next
        'If xString Like "* *" Then
        '    xString = "[" & xString & "]"
        'End If
        If bdOracle Then
            If fld.Type = 135 Then
                l_sql = l_sql & "Chr(39) & FORMAT(X_" & xString & ", " & Chr(34) & "dd/mm/yyyy" & Chr(34) & ") & Chr(39) & " & Chr(34) & ", " & Chr(34)
            'ElseIf fld.Type = 6 Then
                'x_condicao = x_condicao
            ElseIf (fld.Type = 200 Or fld.Type = 129) Then
                l_sql = l_sql & "Chr(39) & X_" & xString & " & Chr(39) & " & Chr(34) & ", " & Chr(34)
            'ElseIf (fld.Type = 200 Or fld.Type = 129) And cbo_operador.Text <> "Semelhante" Then
                'x_condicao = Chr(39) & UCase(x_condicao) & Chr(39)
                'x_ucase = True
            End If
        ElseIf bdAccess Then
            'ACCESS INTEGER
            If fld.Type = vbInteger Then
                xTipoCampo = "sqlNumero"
            ElseIf fld.Type = vbLong Then
                xTipoCampo = "sqlNumero"
            ElseIf fld.Type = vbCurrency Then
                xTipoCampo = "sqlValor"
            ElseIf fld.Type = 135 Then
                xTipoCampo = "sqlData"
            ElseIf fld.Type = 129 Or fld.Type = 200 Then
                xTipoCampo = "sqlTexto"
            ElseIf fld.Type = vbBoolean Then
                xTipoCampo = "sqlBoolean"
            Else
                xTipoCampo = "sqlTexto"
            End If
        ElseIf bdSqlServer Then
            If fld.Type = vbInteger Then
                xTipoCampo = "sqlNumero"
            ElseIf fld.Type = vbLong Then
                xTipoCampo = "sqlNumero"
            ElseIf fld.Type = vbCurrency Then
                xTipoCampo = "sqlValor"
            ElseIf fld.Type = 135 Then
                xTipoCampo = "sqlData"
            ElseIf fld.Type = 129 Or fld.Type = 200 Then
                xTipoCampo = "sqlTexto"
            ElseIf fld.Type = vbBoolean Then
                xTipoCampo = "sqlBoolean"
            Else
                xTipoCampo = "sqlTexto"
            End If
        End If
        
        l_sql = "    Call " & xTipoCampo & "(1, " & "m" & xString & ", " & Chr(34) & ", " & Chr(34) & ")"
        
        i = i + 1
    Next
    If i > 0 Then
        l_sql = Mid(l_sql, 1, (Len(l_sql) - 5)) & Chr(34) & " )" & Chr(34) & ")"
        Print #1, l_sql
    End If
    l_sql = "    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords"
    Print #1, l_sql
    Print #1, "    If lRecordsAffected > 0 Then"
    Print #1, "        Incluir = True"
    Print #1, "    End If"
    Print #1, "    Exit Function"
    Print #1, ""
    Print #1, "Err_Incluir:"
    Print #1, "End Function"
    GeraMetodoIncluir = True
    Exit Function
FileError:
End Function
Function GeraMetodoIncluirNet(ByVal pVbNet2010 As Boolean) As Boolean
Dim i As Integer
Dim i2 As Integer
Dim xString As String
Dim xCampo As String
Dim xTipoCampo As String
Dim xNomeCampo2 As String

On Error GoTo FileError
    GeraMetodoIncluirNet = False
    Print #1, "    Public Function Incluir() As Boolean"
    Print #1, "        Incluir = False"
    Print #1, "        Try"
    If chkAspNet Then
        l_sql = "            sbSQL.Remove(0, sbSQL.Length)"
        Print #1, l_sql
        l_sql = "            sbSQL.Append(" & Chr(34) & "INSERT INTO " & lNomeTabela & " ( "
    Else
        l_sql = "            sbSQL.Clear()"
        Print #1, l_sql
        'l_sql = "            gSQL = " & Chr(34) & "INSERT INTO " & lNomeTabela & " ( "
        l_sql = "            sbSQL.Append(" & Chr(34) & "INSERT INTO " & lNomeTabela & " ( "
    End If
    i = 0
    For Each fld In rst.Fields
        If i = Val(txt_quebra_campo.Text) Then
            'l_sql = l_sql & Chr(34)
            l_sql = l_sql & Chr(34) & ")"
            If chkAspNet Then
                l_sql = l_sql & ")"
            End If
            Print #1, l_sql
            If chkAspNet Then
                l_sql = "            sbSQL.Append(" & Chr(34)
            Else
                'l_sql = "            gSQL += " & Chr(34)
                l_sql = "            sbSQL.Append(" & Chr(34)
            End If
            i = 0
        End If
        xString = fld.name
        If xString Like "* *" Then
            xString = "[" & xString & "]"
        End If
        l_sql = l_sql & xString & ", "
        i = i + 1
    Next
    If i > 0 Then
        l_sql = Mid(l_sql, 1, (Len(l_sql) - 2)) & " ) VALUES ( " & Chr(34)
        l_sql = l_sql & ")"
        If chkAspNet Then
            l_sql = l_sql & ")"
        End If
        Print #1, l_sql
    End If
    l_sql = ""
    i = 0
    For Each fld In rst.Fields
        If i = Val(txt_quebra_variavel.Text) Then
            If pVbNet2010 = True Then
                l_sql = l_sql & Chr(34) & ", " & Chr(34) & ", gBdFuncoesDiversas.bdEnumAzure)"
                Print #1, l_sql
            Else
                Print #1, l_sql
            End If
            If chkAspNet Then
                l_sql = "            sbSQL.Append("
            Else
                l_sql = "            gSQL += "
            End If
            i = 0
        End If
        'xString = fld.Name
        xString = PreparaNomePropriedade(fld.name)
'        For i2 = 1 To Len(fld.name)
'            If Mid(fld.name, i2, 4) = " da " Then
'                i2 = i2 + 3
'            ElseIf Mid(fld.name, i2, 4) = " de " Then
'                i2 = i2 + 3
'            ElseIf Mid(fld.name, i2, 4) = " do " Then
'                i2 = i2 + 3
'            ElseIf Mid(fld.name, i2, 1) <> "_" And Mid(fld.name, i2, 1) <> " " Then
'                xString = xString & Mid(fld.name, i2, 1)
'            End If
'        Next
        'If xString Like "* *" Then
        '    xString = "[" & xString & "]"
        'End If
        xTipoCampo = PreparaTipoCampo2(fld)
        xNomeCampo2 = PreparaNomePropriedade(fld.name)
'        If bdOracle Then
'            If fld.Type = 135 Then
'                l_sql = l_sql & "Chr(39) & FORMAT(X_" & xString & ", " & Chr(34) & "dd/mm/yyyy" & Chr(34) & ") & Chr(39) & " & Chr(34) & ", " & Chr(34)
'            'ElseIf fld.Type = 6 Then
'                'x_condicao = x_condicao
'            ElseIf (fld.Type = 200 Or fld.Type = 129) Then
'                l_sql = l_sql & "Chr(39) & X_" & xString & " & Chr(39) & " & Chr(34) & ", " & Chr(34)
'            'ElseIf (fld.Type = 200 Or fld.Type = 129) And cbo_operador.Text <> "Semelhante" Then
'                'x_condicao = Chr(39) & UCase(x_condicao) & Chr(39)
'                'x_ucase = True
'            End If
'        ElseIf bdAccess Then
'            If fld.Type = vbInteger Then
'                If chkAspNet Then
'                    xTipoCampo = "sqlNumeroSB"
'                Else
'                    xTipoCampo = "sqlNumero"
'                End If
'            ElseIf fld.Type = vbLong Then
'                If chkAspNet Then
'                    xTipoCampo = "sqlNumeroSB"
'                Else
'                    xTipoCampo = "sqlNumero"
'                End If
'            ElseIf fld.Type = vbCurrency Then
'                If chkAspNet Then
'                    xTipoCampo = "sqlValorSB"
'                Else
'                    xTipoCampo = "sqlValor"
'                End If
'            ElseIf fld.Type = 135 Then
'                If chkAspNet Then
'                    xTipoCampo = "sqlDataSB"
'                Else
'                    xTipoCampo = "sqlData"
'                End If
'            ElseIf fld.Type = 129 Or fld.Type = 200 Then
'                If chkAspNet Then
'                    xTipoCampo = "sqlTextoSB"
'                Else
'                    xTipoCampo = "sqlTexto"
'                End If
'            ElseIf fld.Type = vbBoolean Then
'                If chkAspNet Then
'                    xTipoCampo = "sqlBoleanoSB"
'                Else
'                    xTipoCampo = "sqlBoolean"
'                End If
'            Else
'                If chkAspNet Then
'                    xTipoCampo = "sqlTextoSB"
'                Else
'                    xTipoCampo = "sqlTexto"
'                End If
'            End If
'        ElseIf bdSqlServer Then
'            If fld.Type = vbInteger Then
'                If chkAspNet Then
'                    xTipoCampo = "sqlNumeroSB"
'                Else
'                    xTipoCampo = "sqlNumero"
'                End If
'            ElseIf fld.Type = vbLong Then
'                If chkAspNet Then
'                    xTipoCampo = "sqlNumeroSB"
'                Else
'                    xTipoCampo = "sqlNumero"
'                End If
'            ElseIf fld.Type = vbCurrency Then
'                If chkAspNet Then
'                    xTipoCampo = "sqlValorSB"
'                Else
'                    xTipoCampo = "sqlValor"
'                End If
'            ElseIf fld.Type = 135 Then
'                If chkAspNet Then
'                    xTipoCampo = "sqlDataSB"
'                Else
'                    xTipoCampo = "sqlData"
'                End If
'            ElseIf fld.Type = 129 Or fld.Type = 200 Then
'                If chkAspNet Then
'                    xTipoCampo = "sqlTextoSB"
'                Else
'                    xTipoCampo = "sqlTexto"
'                End If
'            ElseIf fld.Type = vbBoolean Then
'                If chkAspNet Then
'                    xTipoCampo = "sqlBoleanoSB"
'                Else
'                    xTipoCampo = "sqlBoolean"
'                End If
'            Else
'                If chkAspNet Then
'                    xTipoCampo = "sqlTextoSB"
'                Else
'                    xTipoCampo = "sqlTexto"
'                End If
'            End If
'        End If
        
        
        If pVbNet2010 = True Then
            If xTipoCampo = "sqlTexto" Then
                l_sql = "            gBdFuncoesDiversas." & xTipoCampo & "Sb(1, " & xNomeCampo2 & ", "
            ElseIf xTipoCampo = "sqlBoolean" Then
                l_sql = "            gBdFuncoesDiversas.sqlBoleanoSb(1, " & xNomeCampo2 & ", "
            Else
                'l_sql = "            " & xString & "(2, " & Chr(34) & xVirgula & xNomeCampo & " = " & Chr(34) & ", " & xNomeCampo2 & ".ToString)"
                l_sql = "            gBdFuncoesDiversas." & xTipoCampo & "Sb(1, " & xNomeCampo2 & ".ToString, "
            End If
        Else
            If xTipoCampo = "sqlTexto" Then
                l_sql = "            " & xTipoCampo & "(1, " & xString & ", " & Chr(34) & ", " & Chr(34) & ")"
            Else
                l_sql = "            " & xTipoCampo & "(1, " & xString & ".ToString, " & Chr(34) & ", " & Chr(34) & ")"
            End If
        End If
        
        
        
        i = i + 1
    Next
    If i > 0 Then
        If pVbNet2010 = True Then
            l_sql = l_sql & Chr(34) & " )" & Chr(34) & ", gBdFuncoesDiversas.bdEnumAzure)"
        Else
            If xTipoCampo = "sqlTexto" Then
                l_sql = Mid(l_sql, 1, (Len(l_sql) - 5)) & Chr(34) & " )" & Chr(34) & ")"
            Else
                l_sql = Mid(l_sql, 1, (Len(l_sql) - 7)) & ".ToString, " & Chr(34) & " )" & Chr(34) & ")"
            End If
        End If
        Print #1, l_sql
    End If
    If chkAspNet Then
        Print #1, "            oleConn.Open()"
        Print #1, "            cmd = New OleDbCommand(sbSQL.ToString, oleConn)"
    Else
        If pVbNet2010 = True Then
        Else
            Print #1, "            cmd = New OleDbCommand(gSQL, gConn)"
        End If
    End If
    If pVbNet2010 = True Then
        Print #1, "            Incluir = gBdFuncoesDiversas.ExecutaCmdAzure(sbSQL.ToString, Me.GetType.Name & " & Chr(34) & ":Incluir" & Chr(34) & ")"
    Else
        Print #1, "            If cmd.ExecuteNonQuery() > 0 Then"
        Print #1, "                Incluir = True"
        Print #1, "            Else"
    End If
    If chkAspNet Then
        Print #1, "                CriaLogRN(Me.GetType.Name & "":Incluir - Erro ao incluir registro."", ""Err.Description"", sbSQL.ToString)"
    Else
        If pVbNet2010 = True Then
        Else
            Print #1, "                CriaLogRN(Me.GetType.Name & "":Incluir - Erro ao incluir registro."", ""Err.Description"", gSQL)"
        End If
    End If
    If pVbNet2010 = True Then
    Else
        Print #1, "            End If"
    End If
    Print #1, "        Catch"
    If chkAspNet Then
        Print #1, "            CriaLogRN(Me.GetType.Name & "":Incluir - Erro não identificado."", Err.Description, sbSQL.ToString)"
    Else
        If pVbNet2010 = True Then
            Print #1, "            gFuncoesDiversas.CriaLog(Me.GetType.Name & "":Incluir - Erro não identificado."", Err.Description, sbSQL.ToString)"
        Else
            Print #1, "            CriaLogRN(Me.GetType.Name & "":Incluir - Erro não identificado."", Err.Description, gSQL)"
            Print #1, "        Finally"
            Print #1, "            cmd.Dispose()"
        End If
    End If
    If chkAspNet Then
        Print #1, "            oleConn.Close()"
    End If
    Print #1, "        End Try"
    Print #1, "    End Function"
    GeraMetodoIncluirNet = True
    Exit Function
FileError:
End Function
Function GeraMetodoLocalizarAnterior() As Boolean
Dim i As Integer
Dim xString As String
Dim xString2 As String

On Error GoTo FileError
    GeraMetodoLocalizarAnterior = False
    Print #1, ""
    Print #1, ""
    Print #1, "Public Function LocalizarAnterior() As Boolean"
    Print #1, "Dim xCondicao As String"
    Print #1, "On Error GoTo trata_erro"
    Print #1, ""
    Print #1, "    LocalizarAnterior = False"
    xString = "    xCondicao = " & Chr(34) & " WHERE "
    For i = 0 To lQtdIndice
        If i > 0 Then
            xString = "    xCondicao = xCondicao & " & Chr(34) & " AND "
        End If
        If lNomeCampoIndice(i) = "Empresa" Then
            xString = xString & lNomeCampoIndiceParentese(i) & " = " & Chr(34) & " & m" & lNomeCampoIndiceAjustado(i)
        Else
            xString2 = "m" & lNomeCampoIndiceAjustado(i)
            If lTipoCampoIndice(i) = "Date" Then
                xString2 = "preparaData(m" & lNomeCampoIndiceAjustado(i) & ")"
            ElseIf lTipoCampoIndice(i) = "String" Then
                xString2 = "preparaTexto(m" & lNomeCampoIndiceAjustado(i) & ")"
            End If
            xString = xString & lNomeCampoIndiceParentese(i) & " < " & Chr(34) & " & " & xString2
            'xString = xString & lNomeCampoIndiceParentese(i) & " < " & Chr(34) & " & m" & lNomeCampoIndiceAjustado(i)
        End If
        Print #1, xString
    Next
    xString = "    gSQL = PreparaSQL(xCondicao, " & Chr(34) & "ORDER BY"
    For i = 0 To lQtdIndice
        If i > 0 Then
            xString = xString & ","
        End If
        xString = xString & " " & lNomeCampoIndiceParentese(i) & " DESC"
    Next
    xString = xString & Chr(34) & ")"
    Print #1, xString
    Print #1, "    LocalizarAnterior = Localizar(1)"
    Print #1, "    If LocalizarAnterior = True Then"
    Print #1, "        AtribuiValor"
    Print #1, "    End If"
    Print #1, "    rs" & lNomeRS & ".Close"
    Print #1, "    Set rs" & lNomeRS & " = Nothing"
    Print #1, "    Exit Function"
    Print #1, ""
    Print #1, "trata_erro:"
    Print #1, "    MsgBox Err.Number & " & Chr(34) & " - " & Chr(34) & " & Err.Description"
    Print #1, "End Function"
    GeraMetodoLocalizarAnterior = True
    Exit Function
FileError:
End Function
Function GeraMetodoLocalizarAnteriorNet(ByVal pVbNet2010 As Boolean) As Boolean
Dim i As Integer
Dim xString As String
Dim xString2 As String

On Error GoTo FileError
    GeraMetodoLocalizarAnteriorNet = False
    Print #1, "    Public Function LocalizarAnterior() As Boolean"
    Print #1, "    Dim xCondicao As String"
    Print #1, ""
    Print #1, "        LocalizarAnterior = False"
    xString = "        xCondicao = " & Chr(34) & " WHERE "
    For i = 0 To lQtdIndice
        If i > 0 Then
            xString = "        xCondicao = xCondicao & " & Chr(34) & " AND "
        End If
        If lNomeCampoIndice(i) = "Empresa" Then
            xString = xString & lNomeCampoIndiceParentese(i) & " = " & Chr(34) & " & m" & lNomeCampoIndiceAjustado(i)
        Else
            xString2 = "m" & lNomeCampoIndiceAjustado(i)
            If lTipoCampoIndice(i) = "Date" Then
                xString2 = "preparaData(m" & lNomeCampoIndiceAjustado(i) & ")"
            ElseIf lTipoCampoIndice(i) = "String" Then
                xString2 = "preparaTexto(m" & lNomeCampoIndiceAjustado(i) & ")"
            End If
            xString = xString & lNomeCampoIndiceParentese(i) & " < " & Chr(34) & " & " & xString2
            'xString = xString & lNomeCampoIndiceParentese(i) & " < " & Chr(34) & " & m" & lNomeCampoIndiceAjustado(i)
        End If
        Print #1, xString
    Next
    If pVbNet2010 Then
        xString = "        PreparaSbSQL(xCondicao, " & Chr(34) & "ORDER BY"
    Else
        xString = "        gSQL = PreparaSQL(xCondicao, " & Chr(34) & "ORDER BY"
    End If
    For i = 0 To lQtdIndice
        If i > 0 Then
            xString = xString & ","
        End If
        xString = xString & " " & lNomeCampoIndiceParentese(i) & " DESC"
    Next
    xString = xString & Chr(34) & ")"
    Print #1, xString
    Print #1, "        LocalizarAnterior = Localizar(1, True, True, True)"
    Print #1, "    End Function"
    GeraMetodoLocalizarAnteriorNet = True
    Exit Function
FileError:
End Function
Function GeraMetodoLocalizarCodigo() As Boolean
Dim i As Integer
Dim xString As String
Dim xString2 As String

On Error GoTo FileError
    GeraMetodoLocalizarCodigo = False
    
    Print #1, ""
    Print #1, ""
    Print #1, "'Inicio Métodos da Classe"
    xString = "Public Function LocalizarCodigo("
    For i = 0 To lQtdIndice
        If i > 0 Then
            xString = xString & ", "
        End If
        xString = xString & "ByVal p" & lNomeCampoIndiceAjustado(i) & " As " & lTipoCampoIndice(i)
    Next
    xString = xString & ") As Boolean"
    Print #1, xString
    Print #1, "Dim xCondicao As String"
    Print #1, "On Error GoTo trata_erro"
    Print #1, ""
    Print #1, "    LocalizarCodigo = False"
    xString = "    xCondicao = " & Chr(34) & " WHERE "
    For i = 0 To lQtdIndice
        If i > 0 Then
            xString = "    xCondicao = xCondicao & " & Chr(34) & " AND "
        End If
        xString2 = "p" & lNomeCampoIndiceAjustado(i)
        If lTipoCampoIndice(i) = "Date" Then
            xString2 = "preparaData(p" & lNomeCampoIndiceAjustado(i) & ")"
        ElseIf lTipoCampoIndice(i) = "String" Then
            xString2 = "preparaTexto(p" & lNomeCampoIndiceAjustado(i) & ")"
        End If
        xString = xString & lNomeCampoIndiceParentese(i) & " = " & Chr(34) & " & " & xString2
        Print #1, xString
    Next
    Print #1, "    gSQL = PreparaSQL(xCondicao, " & Chr(34) & Chr(34) & ")"
    Print #1, "    LocalizarCodigo = Localizar(1)"
    Print #1, "    If LocalizarCodigo = True Then"
    Print #1, "        AtribuiValor"
    Print #1, "    End If"
    Print #1, "    rs" & lNomeRS & ".Close"
    Print #1, "    Set rs" & lNomeRS & " = Nothing"
    Print #1, "    Exit Function"
    Print #1, ""
    Print #1, "trata_erro:"
    Print #1, "    MsgBox Err.Number & " & Chr(34) & " - " & Chr(34) & " & Err.Description"
    Print #1, "End Function"
    GeraMetodoLocalizarCodigo = True
    
    Exit Function
FileError:
End Function
Function GeraMetodoLocalizarCodigoNet(ByVal pVbNet2010 As Boolean) As Boolean
Dim i As Integer
Dim xString As String
Dim xString2 As String

On Error GoTo FileError
    GeraMetodoLocalizarCodigoNet = False
    
    xString = "    Public Function LocalizarCodigo("
    For i = 0 To lQtdIndice
        If i > 0 Then
            xString = xString & ", "
        End If
        xString = xString & "ByVal p" & lNomeCampoIndiceAjustado(i) & " As " & lTipoCampoIndice(i)
    Next
    xString = xString & ") As Boolean"
    Print #1, xString
    Print #1, "        Dim xCondicao As String"
     Print #1, ""
    Print #1, "        LocalizarCodigo = False"
    xString = "        xCondicao = " & Chr(34) & " WHERE "
    For i = 0 To lQtdIndice
        If i > 0 Then
            xString = "        xCondicao = xCondicao & " & Chr(34) & " AND "
        End If
        xString2 = "p" & lNomeCampoIndiceAjustado(i)
        If lTipoCampoIndice(i) = "Date" Then
            xString2 = "preparaData(p" & lNomeCampoIndiceAjustado(i) & ")"
        ElseIf lTipoCampoIndice(i) = "String" Then
            xString2 = "preparaTexto(p" & lNomeCampoIndiceAjustado(i) & ")"
        End If
        xString = xString & lNomeCampoIndiceParentese(i) & " = " & Chr(34) & " & " & xString2
        Print #1, xString
    Next
    If pVbNet2010 Then
        Print #1, "        PreparaSbSQL(xCondicao, " & Chr(34) & Chr(34) & ")"
    Else
        Print #1, "        gSQL = PreparaSQL(xCondicao, " & Chr(34) & Chr(34) & ")"
    End If
    Print #1, "        LocalizarCodigo = Localizar(1, True, True, True)"
    Print #1, "    End Function"
    GeraMetodoLocalizarCodigoNet = True
    
    Exit Function
FileError:
End Function
Function GeraMetodoLocalizarNome() As Boolean
Dim i As Integer
Dim xString As String

On Error GoTo FileError
    GeraMetodoLocalizarNome = False
    Print #1, ""
    Print #1, ""
    Print #1, "Public Function LocalizarNome(ByVal pNome As String) As Boolean"
    Print #1, ""
    Print #1, "On Error GoTo trata_erro"
    Print #1, "    "
    Print #1, "    LocalizarNome = False"
    Print #1, "    gSQL = PreparaSQL(" & Chr(34) & " WHERE Nome = " & Chr(34) & " & preparaTexto(pNome), " & Chr(34) & Chr(34) & ")"
    Print #1, "    LocalizarNome = Localizar(1)"
    Print #1, "    If LocalizarNome = True Then"
    Print #1, "        AtribuiValor"
    Print #1, "    End If"
    Print #1, "    rs" & lNomeRS & ".Close"
    Print #1, "    Set rs" & lNomeRS & " = Nothing"
    Print #1, "    Exit Function"
    Print #1, ""
    Print #1, "trata_erro:"
    Print #1, "    MsgBox Err.Number & " & Chr(34) & " - " & Chr(34) & " & Err.Description"
    Print #1, "End Function"
    GeraMetodoLocalizarNome = True
    Exit Function
FileError:
End Function
Function GeraMetodoLocalizarNomeNet(ByVal pVbNet2010 As Boolean) As Boolean
Dim i As Integer
Dim xString As String

On Error GoTo FileError
    GeraMetodoLocalizarNomeNet = False
    Print #1, "    Public Function LocalizarNome(ByVal pNome As String) As Boolean"
    Print #1, "        LocalizarNome = False"
    If pVbNet2010 Then
        Print #1, "        PreparaSbSQL(" & Chr(34) & " WHERE Nome = " & Chr(34) & " & preparaTexto(pNome), " & Chr(34) & Chr(34) & ")"
    Else
        Print #1, "        gSQL = PreparaSQL(" & Chr(34) & " WHERE Nome = " & Chr(34) & " & preparaTexto(pNome), " & Chr(34) & Chr(34) & ")"
    End If
    Print #1, "        LocalizarNome = Localizar(1, True, True, True)"
    Print #1, "    End Function"
    GeraMetodoLocalizarNomeNet = True
    Exit Function
FileError:
End Function
Function GeraMetodoLocalizarPrimeiro() As Boolean
Dim i As Integer
Dim xString As String
Dim xString2 As String

On Error GoTo FileError
    GeraMetodoLocalizarPrimeiro = False
    Print #1, ""
    Print #1, ""
    Print #1, "Public Function LocalizarPrimeiro() As Boolean"
    Print #1, "Dim xCondicao As String"
    Print #1, "On Error GoTo trata_erro"
    Print #1, ""
    Print #1, "    LocalizarPrimeiro = False"
    xString = "    xCondicao = " & Chr(34) & " WHERE "
    For i = 0 To lQtdIndice
        If i > 0 Then
            xString = "    xCondicao = xCondicao & " & Chr(34) & " AND "
        End If
        If lNomeCampoIndice(i) = "Empresa" Then
            xString = xString & lNomeCampoIndiceParentese(i) & " = " & Chr(34) & " & m" & lNomeCampoIndiceAjustado(i)
        Else
            xString2 = "m" & lNomeCampoIndiceAjustado(i)
            If lTipoCampoIndice(i) = "Date" Then
                xString2 = "preparaData(m" & lNomeCampoIndiceAjustado(i) & ")"
            ElseIf lTipoCampoIndice(i) = "String" Then
                xString2 = "preparaTexto(m" & lNomeCampoIndiceAjustado(i) & ")"
            End If
            xString = xString & lNomeCampoIndiceParentese(i) & " > " & Chr(34) & " & " & xString2
            'xString = xString & lNomeCampoIndiceParentese(i) & " > 0" & Chr(34)
        End If
        Print #1, xString
    Next
    xString = "    gSQL = PreparaSQL(xCondicao, " & Chr(34) & "ORDER BY"
    For i = 0 To lQtdIndice
        If i > 0 Then
            xString = xString & ","
        End If
        xString = xString & " " & lNomeCampoIndiceParentese(i) & " ASC"
    Next
    xString = xString & Chr(34) & ")"
    Print #1, xString
    Print #1, "    LocalizarPrimeiro = Localizar(1)"
    Print #1, "    If LocalizarPrimeiro = True Then"
    Print #1, "        AtribuiValor"
    Print #1, "    End If"
    Print #1, "    rs" & lNomeRS & ".Close"
    Print #1, "    Set rs" & lNomeRS & " = Nothing"
    Print #1, "    Exit Function"
    Print #1, ""
    Print #1, "trata_erro:"
    Print #1, "    MsgBox Err.Number & " & Chr(34) & " - " & Chr(34) & " & Err.Description"
    Print #1, "End Function"
    GeraMetodoLocalizarPrimeiro = True
    Exit Function
FileError:
End Function
Function GeraMetodoLocalizarPrimeiroNet(ByVal pVbNet2010 As Boolean) As Boolean
Dim i As Integer
Dim xString As String
Dim xString2 As String

On Error GoTo FileError
    GeraMetodoLocalizarPrimeiroNet = False
    Print #1, "    Public Function LocalizarPrimeiro() As Boolean"
    Print #1, "    Dim xCondicao As String"
    Print #1, ""
    Print #1, "        LocalizarPrimeiro = False"
    xString = "        xCondicao = " & Chr(34) & " WHERE "
    For i = 0 To lQtdIndice
        If i > 0 Then
            xString = "        xCondicao = xCondicao & " & Chr(34) & " AND "
        End If
        If lNomeCampoIndice(i) = "Empresa" Then
            xString = xString & lNomeCampoIndiceParentese(i) & " = " & Chr(34) & " & m" & lNomeCampoIndiceAjustado(i)
        Else
            xString2 = "m" & lNomeCampoIndiceAjustado(i)
            If lTipoCampoIndice(i) = "Date" Then
                xString2 = "preparaData(m" & lNomeCampoIndiceAjustado(i) & ")"
            ElseIf lTipoCampoIndice(i) = "String" Then
                xString2 = "preparaTexto(m" & lNomeCampoIndiceAjustado(i) & ")"
            End If
            xString = xString & lNomeCampoIndiceParentese(i) & " > " & Chr(34) & " & " & xString2
            'xString = xString & lNomeCampoIndiceParentese(i) & " > 0" & Chr(34)
        End If
        Print #1, xString
    Next
    If pVbNet2010 Then
        xString = "        PreparaSbSQL(xCondicao, " & Chr(34) & "ORDER BY"
    Else
        xString = "        gSQL = PreparaSQL(xCondicao, " & Chr(34) & "ORDER BY"
    End If
    For i = 0 To lQtdIndice
        If i > 0 Then
            xString = xString & ","
        End If
        xString = xString & " " & lNomeCampoIndiceParentese(i) & " ASC"
    Next
    xString = xString & Chr(34) & ")"
    Print #1, xString
    Print #1, "        LocalizarPrimeiro = Localizar(1, True, True, True)"
    Print #1, "    End Function"
    GeraMetodoLocalizarPrimeiroNet = True
    Exit Function
FileError:
End Function
Function GeraMetodoLocalizarProximo() As Boolean
Dim i As Integer
Dim xString As String
Dim xString2 As String

On Error GoTo FileError
    GeraMetodoLocalizarProximo = False
    Print #1, ""
    Print #1, ""
    Print #1, "Public Function LocalizarProximo() As Boolean"
    Print #1, "Dim xCondicao As String"
    Print #1, "On Error GoTo trata_erro"
    Print #1, ""
    Print #1, "    LocalizarProximo = False"
    xString = "    xCondicao = " & Chr(34) & " WHERE "
    For i = 0 To lQtdIndice
        If i > 0 Then
            xString = "    xCondicao = xCondicao & " & Chr(34) & " AND "
        End If
        If lNomeCampoIndice(i) = "Empresa" Then
            xString = xString & lNomeCampoIndiceParentese(i) & " = " & Chr(34) & " & m" & lNomeCampoIndiceAjustado(i)
        Else
            xString2 = "m" & lNomeCampoIndiceAjustado(i)
            If lTipoCampoIndice(i) = "Date" Then
                xString2 = "preparaData(m" & lNomeCampoIndiceAjustado(i) & ")"
            ElseIf lTipoCampoIndice(i) = "String" Then
                xString2 = "preparaTexto(m" & lNomeCampoIndiceAjustado(i) & ")"
            End If
            xString = xString & lNomeCampoIndiceParentese(i) & " > " & Chr(34) & " & " & xString2
            'xString = xString & lNomeCampoIndiceParentese(i) & " > " & Chr(34) & " & m" & lNomeCampoIndiceAjustado(i)
        End If
        Print #1, xString
    Next
    xString = "    gSQL = PreparaSQL(xCondicao, " & Chr(34) & "ORDER BY"
    For i = 0 To lQtdIndice
        If i > 0 Then
            xString = xString & ","
        End If
        xString = xString & " " & lNomeCampoIndiceParentese(i) & " ASC"
    Next
    xString = xString & Chr(34) & ")"
    Print #1, xString
    Print #1, "    LocalizarProximo = Localizar(1)"
    Print #1, "    If LocalizarProximo = True Then"
    Print #1, "        AtribuiValor"
    Print #1, "    End If"
    Print #1, "    rs" & lNomeRS & ".Close"
    Print #1, "    Set rs" & lNomeRS & " = Nothing"
    Print #1, "    Exit Function"
    Print #1, ""
    Print #1, "trata_erro:"
    Print #1, "    MsgBox Err.Number & " & Chr(34) & " - " & Chr(34) & " & Err.Description"
    Print #1, "End Function"
    GeraMetodoLocalizarProximo = True
    Exit Function
FileError:
End Function
Function GeraMetodoLocalizarProximoNet(ByVal pVbNet2010 As Boolean) As Boolean
Dim i As Integer
Dim xString As String
Dim xString2 As String

On Error GoTo FileError
    GeraMetodoLocalizarProximoNet = False
    Print #1, "    Public Function LocalizarProximo() As Boolean"
    Print #1, "    Dim xCondicao As String"
    Print #1, ""
    Print #1, "        LocalizarProximo = False"
    xString = "        xCondicao = " & Chr(34) & " WHERE "
    For i = 0 To lQtdIndice
        If i > 0 Then
            xString = "        xCondicao = xCondicao & " & Chr(34) & " AND "
        End If
        If lNomeCampoIndice(i) = "Empresa" Then
            xString = xString & lNomeCampoIndiceParentese(i) & " = " & Chr(34) & " & m" & lNomeCampoIndiceAjustado(i)
        Else
            xString2 = "m" & lNomeCampoIndiceAjustado(i)
            If lTipoCampoIndice(i) = "Date" Then
                xString2 = "preparaData(m" & lNomeCampoIndiceAjustado(i) & ")"
            ElseIf lTipoCampoIndice(i) = "String" Then
                xString2 = "preparaTexto(m" & lNomeCampoIndiceAjustado(i) & ")"
            End If
            xString = xString & lNomeCampoIndiceParentese(i) & " > " & Chr(34) & " & " & xString2
        End If
        Print #1, xString
    Next
    If pVbNet2010 Then
        xString = "        PreparaSbSQL(xCondicao, " & Chr(34) & "ORDER BY"
    Else
        xString = "        gSQL = PreparaSQL(xCondicao, " & Chr(34) & "ORDER BY"
    End If
    For i = 0 To lQtdIndice
        If i > 0 Then
            xString = xString & ","
        End If
        xString = xString & " " & lNomeCampoIndiceParentese(i) & " ASC"
    Next
    xString = xString & Chr(34) & ")"
    Print #1, xString
    Print #1, "        LocalizarProximo = Localizar(1, True, True, True)"
    Print #1, "    End Function"
    GeraMetodoLocalizarProximoNet = True
    Exit Function
FileError:
End Function
Function GeraMetodoLocalizarUltimo() As Boolean
Dim i As Integer
Dim xString As String
Dim xString2 As String

On Error GoTo FileError
    GeraMetodoLocalizarUltimo = False
    Print #1, ""
    Print #1, ""
    xString = "Public Function LocalizarUltimo("
    For i = 0 To lQtdIndice
        If i > 0 Then
            xString = xString & ", "
        End If
        xString = xString & "ByVal p" & lNomeCampoIndiceAjustado(i) & " As " & lTipoCampoIndice(i)
    Next
    xString = xString & ") As Boolean"
    Print #1, xString
    Print #1, "Dim xCondicao As String"
    Print #1, "On Error GoTo trata_erro"
    Print #1, ""
    Print #1, "    LocalizarUltimo = False"
    xString = "    xCondicao = " & Chr(34) & " WHERE "
    For i = 0 To lQtdIndice
        If i > 0 Then
            xString = "    xCondicao = xCondicao & " & Chr(34) & " AND "
        End If
        If lNomeCampoIndice(i) = "Empresa" Then
            xString = xString & lNomeCampoIndiceParentese(i) & " = " & Chr(34) & " & p" & lNomeCampoIndiceAjustado(i)
        Else
            xString2 = "p" & lNomeCampoIndiceAjustado(i)
            If lTipoCampoIndice(i) = "Date" Then
                xString2 = "preparaData(m" & lNomeCampoIndiceAjustado(i) & ")"
            ElseIf lTipoCampoIndice(i) = "String" Then
                xString2 = "preparaTexto(m" & lNomeCampoIndiceAjustado(i) & ")"
            End If
            xString = xString & lNomeCampoIndiceParentese(i) & " = " & Chr(34) & " & " & xString2
            'xString = xString & lNomeCampoIndiceParentese(i) & " < " & Chr(34) & " & 9999"
        End If
        Print #1, xString
    Next
    xString = "    gSQL = PreparaSQL(xCondicao, " & Chr(34) & "ORDER BY"
    For i = 0 To lQtdIndice
        If i > 0 Then
            xString = xString & ","
        End If
        xString = xString & " " & lNomeCampoIndiceParentese(i) & " DESC"
    Next
    xString = xString & Chr(34) & ")"
    Print #1, xString
    Print #1, "    LocalizarUltimo = Localizar(1)"
    Print #1, "    If LocalizarUltimo = True Then"
    Print #1, "        AtribuiValor"
    Print #1, "    End If"
    Print #1, "    rs" & lNomeRS & ".Close"
    Print #1, "    Set rs" & lNomeRS & " = Nothing"
    Print #1, "    Exit Function"
    Print #1, ""
    Print #1, "trata_erro:"
    Print #1, "    MsgBox Err.Number & " & Chr(34) & " - " & Chr(34) & " & Err.Description"
    Print #1, "End Function"
    GeraMetodoLocalizarUltimo = True
    Exit Function
FileError:
End Function
Function GeraMetodoLocalizarUltimoNet(ByVal pVbNet2010 As Boolean) As Boolean
Dim i As Integer
Dim xString As String
Dim xString2 As String

On Error GoTo FileError
    GeraMetodoLocalizarUltimoNet = False
    xString = "    Public Function LocalizarUltimo("
    For i = 0 To lQtdIndice
        If i > 0 Then
            xString = xString & ", "
        End If
        xString = xString & "ByVal p" & lNomeCampoIndiceAjustado(i) & " As " & lTipoCampoIndice(i)
    Next
    xString = xString & ") As Boolean"
    Print #1, xString
    Print #1, "    Dim xCondicao As String"
    Print #1, ""
    Print #1, "        LocalizarUltimo = False"
    xString = "        xCondicao = " & Chr(34) & " WHERE "
    For i = 0 To lQtdIndice
        If i > 0 Then
            xString = "        xCondicao = xCondicao & " & Chr(34) & " AND "
        End If
        If lNomeCampoIndice(i) = "Empresa" Then
            xString = xString & lNomeCampoIndiceParentese(i) & " = " & Chr(34) & " & p" & lNomeCampoIndiceAjustado(i)
        Else
            xString2 = "p" & lNomeCampoIndiceAjustado(i)
            If lTipoCampoIndice(i) = "Date" Then
                xString2 = "preparaData(m" & lNomeCampoIndiceAjustado(i) & ")"
            ElseIf lTipoCampoIndice(i) = "String" Then
                xString2 = "preparaTexto(m" & lNomeCampoIndiceAjustado(i) & ")"
            End If
            xString = xString & lNomeCampoIndiceParentese(i) & " = " & Chr(34) & " & " & xString2
            'xString = xString & lNomeCampoIndiceParentese(i) & " < " & Chr(34) & " & 9999"
        End If
        Print #1, xString
    Next
    If pVbNet2010 Then
        xString = "        PreparaSbSQL(xCondicao, " & Chr(34) & "ORDER BY"
    Else
        xString = "        gSQL = PreparaSQL(xCondicao, " & Chr(34) & "ORDER BY"
    End If
    For i = 0 To lQtdIndice
        If i > 0 Then
            xString = xString & ","
        End If
        xString = xString & " " & lNomeCampoIndiceParentese(i) & " DESC"
    Next
    xString = xString & Chr(34) & ")"
    Print #1, xString
    Print #1, "        LocalizarUltimo = Localizar(1, True, True, True)"
    Print #1, "    End Function"
    GeraMetodoLocalizarUltimoNet = True
    Exit Function
FileError:
End Function
Function GeraPropriedade() As Boolean
Dim i As Integer
Dim i2 As Integer
Dim xQtdPropriedade As Integer
Dim xNomeCampo As String
Dim xTipoCampo As String

On Error GoTo FileError
    GeraPropriedade = False
    
    Print #1, ""
    Print #1, ""
    Print #1, "'Inicio das Propriedades da Classe"
    
    For Each fld In rst.Fields
        xQtdPropriedade = xQtdPropriedade + 1
        xNomeCampo = ""
        For i2 = 1 To Len(fld.name)
            If Mid(fld.name, i2, 4) = " da " Then
                i2 = i2 + 3
            ElseIf Mid(fld.name, i2, 4) = " de " Then
                i2 = i2 + 3
            ElseIf Mid(fld.name, i2, 4) = " do " Then
                i2 = i2 + 3
            ElseIf Mid(fld.name, i2, 1) <> "_" And Mid(fld.name, i2, 1) <> " " Then
                xNomeCampo = xNomeCampo & Mid(fld.name, i2, 1)
            End If
        Next
        If xQtdPropriedade > 1 Then
            Print #1, ""
            Print #1, ""
        End If
        l_sql = "Public Property Get " & xNomeCampo & "() AS "
        If fld.Type = 131 Then
            If fld.NumericScale = 2 Then
                xTipoCampo = "STRING"
            Else
                If fld.Precision <= 3 Then
                    xTipoCampo = "INTEGER"
                Else
                    xTipoCampo = "LONG"
                End If
            End If
        ElseIf fld.Type = vbInteger Then
            xTipoCampo = "Integer"
        ElseIf fld.Type = vbLong Then
            xTipoCampo = "Long"
        ElseIf fld.Type = vbCurrency Then
            xTipoCampo = "Currency"
        ElseIf fld.Type = vbDate Or fld.Type = 135 Then
            xTipoCampo = "Date"
        ElseIf fld.Type = vbBoolean Then
            xTipoCampo = "Boolean"
        ElseIf (fld.Type = 200 Or fld.Type = 129) Then
            xTipoCampo = "String"
        Else
            xTipoCampo = "String"
        End If
        l_sql = l_sql & xTipoCampo
        Print #1, l_sql
        Print #1, "    " & xNomeCampo & " = m" & xNomeCampo
        Print #1, "End Property"
        Print #1, "Public Property Let " & xNomeCampo & "(ByVal Valor As " & xTipoCampo & ")"
        Print #1, "    m" & xNomeCampo & " = Valor"
        Print #1, "End Property"
    Next
'    Print #1, ""
'    Print #1, ""
'    Print #1, "Public Property Set Conexao(Valor As adodb.Connection)"
'    Print #1, "    Set CConexao = Valor"
'    Print #1, "End Property"
    Print #1, "'Fim das Propriedades da Classe"
    GeraPropriedade = True
    Exit Function
FileError:
End Function
Function GeraPropriedadeNet(ByVal pVbNet2010 As Boolean) As Boolean
Dim i As Integer
Dim i2 As Integer
Dim xQtdPropriedade As Integer
Dim xNomeCampo As String
Dim xTipoCampo As String

On Error GoTo FileError
    GeraPropriedadeNet = False
    
    Print #1, ""
    Print #1, "#Region " & Chr(34) & " Propriedades da Classe " & Chr(34)
    Print #1, ""
    
    For Each fld In rst.Fields
        xQtdPropriedade = xQtdPropriedade + 1
        xNomeCampo = PreparaNomePropriedade(fld.name)
        If pVbNet2010 = True Then
            l_sql = "    Property " & xNomeCampo & "() AS "
        Else
            l_sql = "    Public Property " & xNomeCampo & "() AS "
        End If
        xTipoCampo = PreparaTipoCampo(fld)
'        If fld.Type = 131 And optPostgre.Value = False Then
'            If fld.NumericScale = 2 Then
'                xTipoCampo = "STRING"
'            Else
'                If fld.Precision <= 3 Then
'                    xTipoCampo = "Short"
'                Else
'                    xTipoCampo = "Integer"
'                End If
'            End If
'        ElseIf fld.Type = 131 And optPostgre.Value = True Then
'            If fld.NumericScale = 0 Then
'                If fld.Precision = 39 Then
'                    xTipoCampo = "Decimal"
'                Else
'                    xTipoCampo = "Short"
'                End If
'            End If
'        ElseIf fld.Type = vbInteger Then
'            xTipoCampo = "Short"
'        ElseIf fld.Type = vbLong Or fld.Type = 20 Then
'            xTipoCampo = "Integer"
'        ElseIf fld.Type = vbCurrency Then
'            xTipoCampo = "Decimal"
'        ElseIf fld.Type = vbDate Or fld.Type = 135 Or fld.Type = 133 Then
'            xTipoCampo = "Date"
'        ElseIf fld.Type = vbBoolean Then
'            xTipoCampo = "Boolean"
'        ElseIf (fld.Type = 200 Or fld.Type = 129) Then
'            xTipoCampo = "String"
'        Else
'            xTipoCampo = "String"
'        End If
        l_sql = l_sql & xTipoCampo
        Print #1, l_sql
        If pVbNet2010 = False Then
            Print #1, "        Get"
            'Print #1, "            " & xNomeCampo & " = m" & xNomeCampo
            Print #1, "            Return m" & xNomeCampo
            Print #1, "        End Get"
            Print #1, "        Set" & "(ByVal Value As " & xTipoCampo & ")"
            Print #1, "            m" & xNomeCampo & " = Value"
            Print #1, "        End Set"
            Print #1, "    End Property"
        End If
    Next
    Print #1, ""
    Print #1, "#End Region"
    GeraPropriedadeNet = True
    Exit Function
FileError:
End Function
Function SelecionaGroupBox() As Boolean
Dim xString As String
Dim xInicio As Boolean
Dim xVetor As Variant
Dim i As Integer

On Error GoTo FileError
    SelecionaGroupBox = False
    xInicio = False
    lNomeProgramaFonte = "C:\Cerrado.Net\SGP\" & txt_nome_tabela.Text & ".Vb"
    lNomeGroupBox = ""
    
    If lArqTxt.FileExists(lNomeProgramaFonte) Then
        Set lArquivo = lArqTxt.OpenTextFile(lNomeProgramaFonte, ForReading)
    
        Do Until lArquivo.AtEndOfStream
            xString = lArquivo.ReadLine
            If xInicio = False Then
                If xString Like "*Windows Form Designer*" Then
                    xInicio = True
                End If
            Else
                If xString = "#End Region" Then
                    Exit Do
                End If
                If xString Like "*GroupBox*" Then
                    If xString Like "*Public*" Or xString Like "*Private*" Or xString Like "*Friend*" Then
                        xVetor = Split(xString)
                        For i = LBound(xVetor) To UBound(xVetor)
                            If xVetor(i) = "WithEvents" Then
                                lNomeGroupBox = xVetor(i + 1)
                                If (MsgBox(lNomeGroupBox & Chr(10) & "Este é o GroupBox desejado?", vbQuestion + vbYesNo + vbDefaultButton2, "GroupBox Encontrado!")) = vbYes Then
                                    Exit Do
                                Else
                                    lNomeGroupBox = ""
                                End If
                            End If
                        Next
                    End If
                End If
            End If
        Loop
        lArquivo.Close
        If lNomeGroupBox <> "" Then
            SelecionaGroupBox = True
        End If
    Else
        MsgBox "O programa " & lNomeProgramaFonte & ", não existe!", vbExclamation, "Erro de Verificação"
    End If
    Exit Function

FileError:
End Function
Private Sub LoopGeraEventosObjetos()
    lArquivoDestino.WriteLine ("  ")
    lArquivoDestino.WriteLine ("  ")
    
    rsObjeto.Sort = "Ordem"
    rsObjeto.MoveFirst
    Do Until rsObjeto.EOF
        'MsgBox "Escolha o tipo do campo " & rsObjeto!Nome & ".", vbInformation, "Atenção"
        g_string = "Nome do Campo:" & rsObjeto!Nome & "|@|"
        g_string = g_string & "5|@|"
        g_string = g_string & "1|@|Caixa de Selecao|@|"
        g_string = g_string & "2|@|Caixa de Texto (String)|@|"
        g_string = g_string & "3|@|Caixa de Texto (Inteiro)|@|"
        g_string = g_string & "4|@|Caixa de Texto (Valor)|@|"
        g_string = g_string & "5|@|Caixa de Texto (Data)|@|"
        opcaoGeral.Show 1
        GeraEventosObjetos rsObjeto!Nome, RetiraGString(2)
        rsObjeto.MoveNext
    Loop




    lArquivoDestino.WriteLine ("fim")

End Sub
Function PreparaChavePrimaria() As Boolean
Dim i As Integer
Dim i2 As Integer
Dim xInicio As Integer
Dim xFinal As Integer
Dim xString As String
Dim xTabela As Table
Dim xMindCurrInd As Index
Dim xPadraoIndicesTmp As Object

On Error GoTo FileError
    PreparaChavePrimaria = False
    
    
    'Pega Nome dos Campos da Chave Primária
    Set xTabela = bd_sgp.OpenTable(txt_nome_tabela.Text)
    For Each xPadraoIndicesTmp In xTabela.Indexes
        Set xMindCurrInd = xTabela.Indexes(xPadraoIndicesTmp.name)
        If xMindCurrInd.Primary Then
            lQtdIndice = -1
            xInicio = 0
            xFinal = 0
            xString = xMindCurrInd.Fields
            For i = 1 To Len(xString)
                If Mid(xString, i, 1) = "+" Then
                    lQtdIndice = lQtdIndice + 1
                    xInicio = i + 1
                End If
                If Mid(xString, i, 1) = ";" Then
                    xFinal = i - 1
                End If
                If xInicio > 0 And xFinal > 0 Then
                    lNomeCampoIndice(lQtdIndice) = Mid(xString, xInicio, (xFinal - xInicio) + 1)
                    xInicio = 0
                    xFinal = 0
                End If
            Next
        End If
    Next
    xTabela.Close
    lNomeCampoIndice(lQtdIndice) = Mid(xString, xInicio, (Len(xString) - xInicio) + 1)
    
    For i = 0 To lQtdIndice
        lNomeCampoIndiceAjustado(i) = ""
        For i2 = 1 To Len(lNomeCampoIndice(i))
            If Mid(lNomeCampoIndice(i), i2, 4) = " da " Then
                i2 = i2 + 3
            ElseIf Mid(lNomeCampoIndice(i), i2, 4) = " de " Then
                i2 = i2 + 3
            ElseIf Mid(lNomeCampoIndice(i), i2, 4) = " do " Then
                i2 = i2 + 3
            ElseIf Mid(lNomeCampoIndice(i), i2, 1) <> "_" And Mid(lNomeCampoIndice(i), i2, 1) <> " " Then
                lNomeCampoIndiceAjustado(i) = lNomeCampoIndiceAjustado(i) & Mid(lNomeCampoIndice(i), i2, 1)
            End If
        Next
    Next
    
    For i = 0 To lQtdIndice
        lNomeCampoIndiceParentese(i) = lNomeCampoIndice(i)
        If lNomeCampoIndiceParentese(i) Like "* *" Then
            lNomeCampoIndiceParentese(i) = "[" & lNomeCampoIndiceParentese(i) & "]"
        End If
    Next
    
    
    'Pega Tipo dos Campos da Chave Primária
    For i = 0 To lQtdIndice
        For Each fld In rst.Fields
            If fld.name = lNomeCampoIndice(i) Then
                If fld.Type = 131 Then
                    If fld.NumericScale = 2 Then
                        lTipoCampoIndice(i) = "String"
                    Else
                        If fld.Precision <= 3 Then
                            If chkVbNet.Value = 1 Then
                                lTipoCampoIndice(i) = "Short"
                            Else
                                lTipoCampoIndice(i) = "Integer"
                            End If
                        Else
                            If chkVbNet.Value = 1 Then
                                lTipoCampoIndice(i) = "Integer"
                            Else
                                lTipoCampoIndice(i) = "Long"
                            End If
                        End If
                    End If
                ElseIf fld.Type = vbInteger Then
                    If chkVbNet.Value = 1 Then
                        lTipoCampoIndice(i) = "Short"
                    Else
                        lTipoCampoIndice(i) = "Integer"
                    End If
                ElseIf fld.Type = vbLong Then
                    If chkVbNet.Value = 1 Then
                        lTipoCampoIndice(i) = "Integer"
                    Else
                        lTipoCampoIndice(i) = "Long"
                    End If
                ElseIf fld.Type = vbCurrency Then
                    If chkVbNet.Value = 1 Then
                        lTipoCampoIndice(i) = "Decimal"
                    Else
                        lTipoCampoIndice(i) = "Currency"
                    End If
                ElseIf fld.Type = vbDate Or fld.Type = 135 Then
                    lTipoCampoIndice(i) = "Date"
                ElseIf (fld.Type = 200 Or fld.Type = 129) Then
                    lTipoCampoIndice(i) = "String"
                Else
                    lTipoCampoIndice(i) = "String"
                End If
                Exit For
            End If
        Next
    Next
    PreparaChavePrimaria = True
    Exit Function
FileError:
End Function
Function PreparaNomePropriedade(ByVal pNomeCampo As String) As String
    Dim i As Integer
    Dim xString As String
    Dim xMaiusculo As Boolean
    
    xString = ""
    If chkVbNet2010.Value = 1 And optPostgre.Value = True Then
        xMaiusculo = True
    End If
    For i = 1 To Len(pNomeCampo)
        If Mid(pNomeCampo, i, 4) = " da " Or Mid(pNomeCampo, i, 4) = "_da_" Then
            i = i + 3
            If chkVbNet2010.Value = 1 And optPostgre.Value = True Then
                xMaiusculo = True
            End If
        ElseIf Mid(pNomeCampo, i, 4) = " de " Or Mid(pNomeCampo, i, 4) = "_de_" Then
            i = i + 3
            If chkVbNet2010.Value = 1 And optPostgre.Value = True Then
                xMaiusculo = True
            End If
        ElseIf Mid(pNomeCampo, i, 4) = " do " Or Mid(pNomeCampo, i, 4) = "_do_" Then
            i = i + 3
            If chkVbNet2010.Value = 1 And optPostgre.Value = True Then
                xMaiusculo = True
            End If
        ElseIf Mid(pNomeCampo, i, 1) = "_" Then
            If chkVbNet2010.Value = 1 And optPostgre.Value = True Then
                xMaiusculo = True
            End If
        ElseIf Mid(pNomeCampo, i, 1) <> "_" And Mid(pNomeCampo, i, 1) <> " " Then
            If xMaiusculo = True Then
                xMaiusculo = False
                xString = xString & UCase(Mid(pNomeCampo, i, 1))
            Else
                xString = xString & Mid(pNomeCampo, i, 1)
            End If
        End If
    Next
    PreparaNomePropriedade = xString
End Function
Private Function PreparaTipoCampo(ByVal pCampo As adodb.Field) As String
    Dim xTipoCampo As String
    
    xTipoCampo = ""
    If pCampo.Type = 131 And optPostgre.Value = False Then
        If pCampo.NumericScale = 2 Then
            xTipoCampo = "STRING"
        Else
            If pCampo.Precision <= 3 Then
                xTipoCampo = "Short"
            Else
                xTipoCampo = "Integer"
            End If
        End If
    ElseIf pCampo.Type = 131 And optPostgre.Value = True Then
        If pCampo.NumericScale = 0 Then
            If pCampo.Precision = 39 Then
                xTipoCampo = "Decimal"
            Else
                xTipoCampo = "Short"
            End If
        End If
    ElseIf pCampo.Type = vbInteger Then
        xTipoCampo = "Short"
    ElseIf pCampo.Type = vbLong Or pCampo.Type = 20 Then
        xTipoCampo = "Integer"
    ElseIf pCampo.Type = vbCurrency Then
        xTipoCampo = "Decimal"
    ElseIf pCampo.Type = vbDate Or pCampo.Type = 135 Or pCampo.Type = 133 Then
        xTipoCampo = "Date"
    ElseIf pCampo.Type = vbBoolean Then
        xTipoCampo = "Boolean"
    ElseIf (pCampo.Type = 200 Or pCampo.Type = 129) Then
        xTipoCampo = "String"
    Else
        xTipoCampo = "String"
    End If
    PreparaTipoCampo = xTipoCampo
End Function
Private Function PreparaTipoCampo2(ByVal pCampo As adodb.Field) As String
    Dim xTipoCampo As String
    
    xTipoCampo = ""
    If bdOracle Then
    ElseIf optPostgre.Value = True Then
        If pCampo.Type = vbInteger Then
            If chkAspNet Then
                xTipoCampo = "sqlNumeroSB"
            Else
                xTipoCampo = "sqlNumero"
            End If
        ElseIf pCampo.Type = vbLong Or pCampo.Type = 20 Then
            If chkAspNet Then
                xTipoCampo = "sqlNumeroSB"
            Else
                xTipoCampo = "sqlNumero"
            End If
        ElseIf pCampo.Type = vbCurrency Or pCampo.Type = 131 Then
            If chkAspNet Then
                xTipoCampo = "sqlValorSB"
            Else
                xTipoCampo = "sqlValor"
            End If
        ElseIf pCampo.Type = 135 Or pCampo.Type = 133 Then
            If chkAspNet Then
                xTipoCampo = "sqlDataSB"
            Else
                xTipoCampo = "sqlData"
            End If
        ElseIf pCampo.Type = 129 Or pCampo.Type = 200 Or pCampo.Type = 202 Then
            If chkAspNet Then
                xTipoCampo = "sqlTextoSB"
            Else
                xTipoCampo = "sqlTexto"
            End If
        ElseIf pCampo.Type = vbBoolean Then
            If chkAspNet Then
                xTipoCampo = "sqlBoleanoSB"
            Else
                xTipoCampo = "sqlBoolean"
            End If
        Else
            If chkAspNet Then
                xTipoCampo = "sqlTextoSB"
            Else
                xTipoCampo = "sqlTexto"
            End If
        End If
    ElseIf bdAccess Then
        If pCampo.Type = vbInteger Then
            If chkAspNet Then
                xTipoCampo = "sqlNumeroSB"
            Else
                xTipoCampo = "sqlNumero"
            End If
        ElseIf pCampo.Type = vbLong Then
            If chkAspNet Then
                xTipoCampo = "sqlNumeroSB"
            Else
                xTipoCampo = "sqlNumero"
            End If
        ElseIf pCampo.Type = vbCurrency Then
            If chkAspNet Then
                xTipoCampo = "sqlValorSB"
            Else
                xTipoCampo = "sqlValor"
            End If
        ElseIf pCampo.Type = 135 Then
            If chkAspNet Then
                xTipoCampo = "sqlDataSB"
            Else
                xTipoCampo = "sqlData"
            End If
        ElseIf pCampo.Type = 129 Or pCampo.Type = 200 Then
            If chkAspNet Then
                xTipoCampo = "sqlTextoSB"
            Else
                xTipoCampo = "sqlTexto"
            End If
        ElseIf pCampo.Type = vbBoolean Then
            If chkAspNet Then
                xTipoCampo = "sqlBoleanoSB"
            Else
                xTipoCampo = "sqlBoolean"
            End If
        Else
            xTipoCampo = "sqlTexto"
        End If
    ElseIf bdSqlServer Then
        If pCampo.Type = vbInteger Then
            If chkAspNet Then
                xTipoCampo = "sqlNumeroSB"
            Else
                xTipoCampo = "sqlNumero"
            End If
        ElseIf pCampo.Type = vbLong Then
            If chkAspNet Then
                xTipoCampo = "sqlNumeroSB"
            Else
                xTipoCampo = "sqlNumero"
            End If
        ElseIf pCampo.Type = vbCurrency Then
            If chkAspNet Then
                xTipoCampo = "sqlValorSB"
            Else
                xTipoCampo = "sqlValor"
            End If
        ElseIf pCampo.Type = 135 Then
            If chkAspNet Then
                xTipoCampo = "sqlDataSB"
            Else
                xTipoCampo = "sqlData"
            End If
        ElseIf pCampo.Type = 129 Or pCampo.Type = 200 Then
            If chkAspNet Then
                xTipoCampo = "sqlTextoSB"
            Else
                xTipoCampo = "sqlTexto"
            End If
        ElseIf pCampo.Type = vbBoolean Then
            If chkAspNet Then
                xTipoCampo = "sqlBoleanoSB"
            Else
                xTipoCampo = "sqlBoolean"
            End If
        Else
            If chkAspNet Then
                xTipoCampo = "sqlTextoSB"
            Else
                xTipoCampo = "sqlTexto"
            End If
        End If
    End If
    PreparaTipoCampo2 = xTipoCampo
End Function
Private Function ProximoObjeto(ByVal pOrdemObjeto As String) As String
    ProximoObjeto = "btnOK"

    rsObjeto.Sort = "Ordem"
    rsObjeto.MoveFirst
    rsObjeto.Find "Ordem='" & pOrdemObjeto & "'"
    If rsObjeto.EOF = False Then
        rsObjeto.MoveNext
        If rsObjeto.EOF = False Then
            ProximoObjeto = rsObjeto!Nome
        End If
    End If
    rsObjeto.MoveFirst
    rsObjeto.Find "Ordem='" & pOrdemObjeto & "'"
End Function
Private Sub cmd_ok_Click()
    Dim retval As Long
    Dim xStringConexao As String
    
    
    If optDadosInternet.Value = True Then
        cnnSGPDados.Mode = adModeRead
        Set cnnSGPDados = New adodb.Connection
        'cnnSGPDados.Open "Driver={Microsoft Access Driver (*.mdb)};Dbq=C:\Inetpub\wwwroot\login\Dados\dados.mdb;Uid=Admin;Pwd=;"
        'cnnSGPDados.Open "Driver={Microsoft Access Driver (*.mdb)};Dbq=C:\Inetpub\wwwroot\login\Dados\dados.mdb;Uid=Admin;Pwd=;"
        cnnSGPDados.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=C:\Inetpub\wwwroot\login\Dados\Acesso.mdb"
        cnnSGPDados.Open
    ElseIf optSgleData.Value = True Then
        cnnSGPDados.Mode = adModeRead
        Set cnnSGPDados = New adodb.Connection
        cnnSGPDados.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=C:\Cerrado\Sgle\Data\Sgle_data.mdb"
        cnnSGPDados.Open
    ElseIf optComercial.Value = True Then
        cnnSGPDados.Mode = adModeRead
        Set cnnSGPDados = New adodb.Connection
        cnnSGPDados.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=C:\Inetpub\wwwroot\Comercial\Comercial\Dados\Comercial.mdb"
        cnnSGPDados.Open
    ElseIf optOutroBancoAccess.Value = True Then
        cnnSGPDados.Mode = adModeRead
        Set cnnSGPDados = New adodb.Connection
        cnnSGPDados.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=" & txtNomeBancoAccess.Text
        cnnSGPDados.Open
    End If
    
    
    If chkGeraEventos.Value = 1 Then
        If SelecionaGroupBox Then
            Set lArquivoDestino = lArqTxt.CreateTextFile("\VB5\SGP\DATA\modulo_de_classe.txt")
            If CarregaObjetos Then
                AtribuiOrdemObjetos
                LoopGeraEventosObjetos
            End If
            lArquivoDestino.Close
            retval = Shell("c:\WINDOWS\NOTEPAD.EXE \VB5\SGP\DATA\modulo_de_classe.txt", 1)
        End If
    Else
        If opt_sgp_data.Value = True Then
            xStringConexao = Conectar.ConnectionString
        ElseIf optNuvemNFe.Value = True Then
            'xStringConexao = "PROVIDER=SQLOLEDB.1;DATA SOURCE=" & "127.0.0.1,4949" & ";INITIAL CATALOG=" & "NuvemNFe" & ";USER ID=sa;PASSWORD=" & gSenhaBD & ";"
            xStringConexao = "PROVIDER=SQLOLEDB.1;DATA SOURCE=" & "127.0.0.1" & gPortaBanco & ";INITIAL CATALOG=" & "NuvemNFe" & ";USER ID=sa;PASSWORD=" & gSenhaBD & ";"
        ElseIf opt_sgc_data.Value = True Then
            'xStringConexao = "PROVIDER=SQLOLEDB.1;DATA SOURCE=" & "127.0.0.1,4949" & ";INITIAL CATALOG=" & "sgc_data" & ";USER ID=sa;PASSWORD=" & gSenhaBD & ";"
            xStringConexao = "PROVIDER=SQLOLEDB.1;DATA SOURCE=" & "127.0.0.1" & gPortaBanco & ";INITIAL CATALOG=" & "sgc_data" & ";USER ID=sa;PASSWORD=" & gSenhaBD & ";"
        ElseIf opt_sfa_data.Value = True Then
            'xStringConexao = "PROVIDER=SQLOLEDB.1;DATA SOURCE=" & "127.0.0.1,4949" & ";INITIAL CATALOG=" & "Sfa_Data" & ";USER ID=sa;PASSWORD=" & gSenhaBD & ";"
            xStringConexao = "PROVIDER=SQLOLEDB.1;DATA SOURCE=" & "127.0.0.1" & gPortaBanco & ";INITIAL CATALOG=" & "Sfa_Data" & ";USER ID=sa;PASSWORD=" & gSenhaBD & ";"
        ElseIf optDadosInternet.Value = True Or optSgleData.Value = True Or optComercial.Value = True Or optOutroBancoAccess.Value = True Then
            xStringConexao = cnnSGPDados.ConnectionString
        ElseIf optCerradoData.Value = True Then
            'xStringConexao = "PROVIDER=SQLOLEDB.1;DATA SOURCE=" & "127.0.0.1,4949" & ";INITIAL CATALOG=" & "CerradoData" & ";USER ID=sa;PASSWORD=" & gSenhaBD & ";"
            xStringConexao = "PROVIDER=SQLOLEDB.1;DATA SOURCE=" & "127.0.0.1" & gPortaBanco & ";INITIAL CATALOG=" & "CerradoData" & ";USER ID=sa;PASSWORD=" & gSenhaBD & ";"
        ElseIf optTefCerrado.Value = True Then
            xStringConexao = "PROVIDER=SQLOLEDB.1;DATA SOURCE=" & "192.168.1.77,4848" & ";INITIAL CATALOG=" & "TefCerrado" & ";USER ID=sa;PASSWORD=" & gSenhaBD & ";"
            'xStringConexao = "PROVIDER=SQLOLEDB.1;DATA SOURCE=" & "127.0.0.1,4949" & ";INITIAL CATALOG=" & "TefCerrado" & ";USER ID=sa;PASSWORD=" & gSenhaBD & ";"
        ElseIf optGateData.Value = True Then
            'xStringConexao = "PROVIDER=SQLOLEDB.1;DATA SOURCE=" & "127.0.0.1,4949" & ";INITIAL CATALOG=" & "GateData" & ";USER ID=sa;PASSWORD=" & gSenhaBD & ";"
            xStringConexao = "PROVIDER=SQLOLEDB.1;DATA SOURCE=" & "127.0.0.1" & gPortaBanco & ";INITIAL CATALOG=" & "GateData" & ";USER ID=sa;PASSWORD=" & gSenhaBD & ";"
        ElseIf optPostgre.Value = True Then
            xStringConexao = "Provider=PostgreSQL OLE DB Provider;Data Source=" & "127.0.0.1" & ";location=" & txtNomeBancoAccess.Text & ";User ID=postgres;password=" & gSenhaBD & ";"
            xStringConexao = "Provider=PostgreSQL OLE DB Provider;Data Source=" & "192.168.1.194" & ";location=" & txtNomeBancoAccess.Text & ";User ID=postgres;password=" & gSenhaBD & ";"
        End If
        If ConexaoAuxiliar.AbreConexao(xStringConexao) = False Then
            MsgBox "Erro ao conectar ao banco de dados: " & xStringConexao, vbCritical, "Erro de Conexão!"
            Exit Sub
        End If
        AtualizaRecordset
        PreparaChavePrimaria
        If CriaArquivoTexto Then
            If chkVbNet.Value = 1 Or chkVbNet2010.Value = 1 Then
                GeraDeclaracaoNet (chkVbNet2010.Value)
                GeraPropriedadeNet (chkVbNet2010.Value)
                GeraMetodoAlterarNet (chkVbNet2010.Value)
                GeraMetodoExcluirNet
                GeraMetodoIncluirNet (chkVbNet2010.Value)
                GeraMetodoLocalizarAnteriorNet (chkVbNet2010.Value)
                GeraMetodoLocalizarCodigoNet (chkVbNet2010.Value)
                GeraMetodoLocalizarNomeNet (chkVbNet2010.Value)
                GeraMetodoLocalizarPrimeiroNet (chkVbNet2010.Value)
                GeraMetodoLocalizarProximoNet (chkVbNet2010.Value)
                GeraMetodoLocalizarUltimoNet (chkVbNet2010.Value)
                GeraFuncoesInternasNet (chkVbNet2010.Value)
            Else
                GeraDeclaracao
                GeraPropriedade
                GeraMetodoLocalizarCodigo
                GeraMetodoLocalizarAnterior
                GeraMetodoLocalizarNome
                GeraMetodoLocalizarPrimeiro
                GeraMetodoLocalizarProximo
                GeraMetodoLocalizarUltimo
                GeraMetodoIncluir
                GeraMetodoAlterar
                GeraMetodoExcluir
                GeraFuncoesInternas
                'GeraArquivoStringInsert
            End If
            FechaArquivoTexto
            rst.Close
            Set rst = Nothing
        End If
    End If
    
    If optDadosInternet.Value = True Or optSgleData.Value = True Or optComercial.Value = True Or optOutroBancoAccess.Value = True Then
        cnnSGPDados.Close
    End If
End Sub
Private Sub cmd_ok_KeyPress(KeyAscii As Integer)
    If KeyAscii = 9 Then 'Crtl + I
        zzImportaCupomFiscal
    End If
End Sub

Private Sub Form_Load()
    Set ConexaoAuxiliar = New cConexaoAuxiliar
    txt_quebra_campo.Text = "5"
    txt_quebra_variavel.Text = "1"
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub zzImportaCupomFiscal()
    Dim xSQL As String
    Dim rstMovCupomFiscalCab As adodb.Recordset
    Dim rstMovCupomFiscalItem As adodb.Recordset
    Dim rstAuxiliar As adodb.Recordset
    Dim MovimentoCupomFiscal As New cMovimentoCupomFiscal
    Dim MovimentoCupomFiscalItem As New cMovimentoCupomFiscalItem
    Dim Produto As New cProduto
    Dim xCodigoProduto As Integer
    Dim xFormaPagamento As Integer

    cnnSGPDados.Mode = adModeRead
    Set cnnSGPDados = New adodb.Connection
    cnnSGPDados.ConnectionString = "Driver={Microsoft Access Driver (*.mdb)};Dbq=C:\Backup_SGP\Fita detalhe vera cruz 01 09 11\Fita detalhe\teste.mdb;Uid=Admin;Pwd=;"
    'cnnSGPDados.Open "Driver={Microsoft Access Driver (*.mdb)};Dbq=C:\Inetpub\wwwroot\login\Dados\dados.mdb;Uid=Admin;Pwd=;"
    'cnnSGPDados.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=C:\Inetpub\wwwroot\login\Dados\Acesso.mdb"
    cnnSGPDados.Open
    
    xSQL = ""
    xSQL = xSQL & "SELECT *"
    xSQL = xSQL & "  FROM CuponsFiscais"
    xSQL = xSQL & " WHERE contCOO < " & 149339
    xSQL = xSQL & " ORDER BY contCOO"
    Set rstMovCupomFiscalCab = New adodb.Recordset
    rstMovCupomFiscalCab.CursorLocation = adUseClient 'adUseServer  'adUseClient
    rstMovCupomFiscalCab.Open xSQL, cnnSGPDados, adOpenForwardOnly, adLockReadOnly
    
    
    
    With rstMovCupomFiscalCab
        If .RecordCount > 0 Then
            Do Until .EOF
                If !HoraInicio >= CDate("01/07/2011") Then
                    xFormaPagamento = 1
                    
                    'Busca Forma de Pagamento
                    xSQL = ""
                    xSQL = xSQL & "SELECT IDTipo"
                    xSQL = xSQL & "  FROM PgtosPorDocumento"
                    xSQL = xSQL & " WHERE contCOODoc = " & !contCOO
                    xSQL = xSQL & " ORDER BY IDTipo"
                    Set rstAuxiliar = New adodb.Recordset
                    rstAuxiliar.CursorLocation = adUseClient 'adUseServer  'adUseClient
                    rstAuxiliar.Open xSQL, cnnSGPDados, adOpenForwardOnly, adLockReadOnly
                    If rstAuxiliar.RecordCount > 0 Then
                        rstAuxiliar.MoveFirst
                        If rstAuxiliar!IDTipo = 1 Then
                            xFormaPagamento = 1
                        ElseIf rstAuxiliar!IDTipo = 2 Or rstAuxiliar!IDTipo = 5 Or rstAuxiliar!IDTipo = 11 Then
                            xFormaPagamento = 2
                        ElseIf rstAuxiliar!IDTipo = 3 Or rstAuxiliar!IDTipo = 7 Then
                            xFormaPagamento = 3
                        ElseIf rstAuxiliar!IDTipo = 4 Or rstAuxiliar!IDTipo = 16 Or rstAuxiliar!IDTipo = 17 Or rstAuxiliar!IDTipo = 18 Or rstAuxiliar!IDTipo = 19 Then
                            xFormaPagamento = 4
                        ElseIf rstAuxiliar!IDTipo = 6 Or rstAuxiliar!IDTipo = 9 Or rstAuxiliar!IDTipo = 15 Then
                            xFormaPagamento = 5
                        Else
                            xFormaPagamento = 4
                        End If
                    End If
                    
                
                    xSQL = ""
                    xSQL = xSQL & "SELECT *"
                    xSQL = xSQL & "  FROM ItensPorCupom"
                    xSQL = xSQL & " WHERE contCOOCupom = " & !contCOO
                    xSQL = xSQL & " ORDER BY contCOOCupom, Ordem"
                    Set rstMovCupomFiscalItem = New adodb.Recordset
                    rstMovCupomFiscalItem.CursorLocation = adUseClient 'adUseServer  'adUseClient
                    rstMovCupomFiscalItem.Open xSQL, cnnSGPDados, adOpenForwardOnly, adLockReadOnly
                    If rstMovCupomFiscalItem.RecordCount > 0 Then
                        Do Until rstMovCupomFiscalItem.EOF
                            
                            
                            xCodigoProduto = 0
                            'Converte Codigo do Produto
                            xSQL = ""
                            xSQL = xSQL & "SELECT ['Codigo SGP'] AS CodigoSgp"
                            xSQL = xSQL & "  FROM ItensCadastrados"
                            xSQL = xSQL & " WHERE IDCodigoItem = " & rstMovCupomFiscalItem!CodigoItem
                            Set rstAuxiliar = New adodb.Recordset
                            rstAuxiliar.CursorLocation = adUseClient 'adUseServer  'adUseClient
                            rstAuxiliar.Open xSQL, cnnSGPDados, adOpenForwardOnly, adLockReadOnly
                            If rstAuxiliar.RecordCount > 0 Then
                                rstAuxiliar.MoveFirst
                                xCodigoProduto = rstAuxiliar!CodigoSgp
                            Else
                                MsgBox "Nao foi possivel converter o codigo do produto"
                            End If
                            
                            
                            If Produto.LocalizarCodigo(xCodigoProduto) Then
                                MovimentoCupomFiscal.Empresa = g_empresa
                                MovimentoCupomFiscal.NumeroCupom = !contCOO
                                MovimentoCupomFiscal.Ordem = rstMovCupomFiscalItem!Ordem
                                MovimentoCupomFiscal.Data = Format(!HoraInicio, "dd/mm/yyyy")
                                MovimentoCupomFiscal.Hora = Format(!HoraInicio, "hh:mm:ss")
                                MovimentoCupomFiscal.DataCupom = Format(!HoraInicio, "dd/mm/yyyy")
                                MovimentoCupomFiscal.Periodo = 1
                                MovimentoCupomFiscal.TipoMovimento = 2
                                MovimentoCupomFiscal.CodigoCliente = 0
                                MovimentoCupomFiscal.CodigoProduto = xCodigoProduto
                                MovimentoCupomFiscal.ValorUnitario = rstMovCupomFiscalItem!ValorUnitario
                                MovimentoCupomFiscal.Quantidade = rstMovCupomFiscalItem!Quantidade
                                MovimentoCupomFiscal.ValorTotal = rstMovCupomFiscalItem!ValorTotal
                                MovimentoCupomFiscal.FormaPagamento = xFormaPagamento
                                MovimentoCupomFiscal.ValorRecebido = 0
                                If Not IsNull(!total) Then
                                    MovimentoCupomFiscal.ValorRecebido = !total
                                End If
                                MovimentoCupomFiscal.NumeroCheque = ""
                                MovimentoCupomFiscal.Telefone = ""
                                MovimentoCupomFiscal.operador = 1
                                MovimentoCupomFiscal.CupomCancelado = !Cancelado
                                MovimentoCupomFiscal.ItemCancelado = rstMovCupomFiscalItem!Cancelado
                                'FF
                                If rstMovCupomFiscalItem!SituacaoTributaria = 1 Then
                                    MovimentoCupomFiscal.CodigoAliquota = 2
                                'II
                                ElseIf rstMovCupomFiscalItem!SituacaoTributaria = 2 Then
                                    MovimentoCupomFiscal.CodigoAliquota = 1
                                'NN
                                ElseIf rstMovCupomFiscalItem!SituacaoTributaria = 3 Then
                                    MovimentoCupomFiscal.CodigoAliquota = 3
                                '17%
                                ElseIf rstMovCupomFiscalItem!SituacaoTributaria = 7 Then
                                    MovimentoCupomFiscal.CodigoAliquota = 5
                                Else
                                    MsgBox "Aliquota?"
                                End If
                                MovimentoCupomFiscal.ValorDesconto = 0
                                If !ValorDesconto > 0 Then
                                    MovimentoCupomFiscal.ValorDesconto = !ValorDesconto
                                ElseIf !ValorAcrescimo > 0 Then
                                    MovimentoCupomFiscal.ValorDesconto = -!ValorAcrescimo
                                End If
                                MovimentoCupomFiscal.Nome = ""
                                MovimentoCupomFiscal.CPFCNPJ = ""
                                If Mid(![CNPJ_CPFConsumidor], 1, 11) = "11111111111" Then
                                    MovimentoCupomFiscal.CPFCNPJ = ""
                                End If
                                MovimentoCupomFiscal.TipoCombustivel = Produto.TipoCombustivel
                                MovimentoCupomFiscal.CodigoECF = 1
                                MovimentoCupomFiscal.CodigoGrupo = Produto.CodigoGrupo
                                MovimentoCupomFiscal.TipoSubEstoque = 2
                                MovimentoCupomFiscal.ValorDescontoEmbutido = 0
                                If Not MovimentoCupomFiscal.Incluir Then
                                    MsgBox "Não foi possível incluir o cupom fiscal.", vbInformation, "Erro de Integridade!"
                                End If
                                
                                
                                MovimentoCupomFiscalItem.Empresa = g_empresa
                                MovimentoCupomFiscalItem.NumeroCupom = !contCOO
                                MovimentoCupomFiscalItem.Ordem = rstMovCupomFiscalItem!Ordem
                                MovimentoCupomFiscalItem.Data = Format(!HoraInicio, "dd/mm/yyyy")
                                MovimentoCupomFiscalItem.CodigoProduto = xCodigoProduto
                                MovimentoCupomFiscalItem.ValorUnitario = rstMovCupomFiscalItem!ValorUnitario
                                MovimentoCupomFiscalItem.Quantidade = rstMovCupomFiscalItem!Quantidade
                                MovimentoCupomFiscalItem.ValorTotal = rstMovCupomFiscalItem!ValorTotal
                                MovimentoCupomFiscalItem.ItemCancelado = rstMovCupomFiscalItem!Cancelado
                                MovimentoCupomFiscalItem.ValorDesconto = 0
                                MovimentoCupomFiscalItem.ValorAcrescimo = 0
                                If rstMovCupomFiscalItem!ValorDesconto > 0 Then
                                    MovimentoCupomFiscalItem.ValorDesconto = rstMovCupomFiscalItem!ValorDesconto
                                End If
                                If rstMovCupomFiscalItem!ValorAcrescimo > 0 Then
                                    MovimentoCupomFiscalItem.ValorAcrescimo = rstMovCupomFiscalItem!ValorAcrescimo
                                End If
                                MovimentoCupomFiscalItem.DescontoEmbutido = False
                                MovimentoCupomFiscalItem.Periodo = 1
                                MovimentoCupomFiscalItem.TipoCombustivel = Produto.TipoCombustivel
                                MovimentoCupomFiscalItem.CodigoECF = 1
                                'FF
                                If rstMovCupomFiscalItem!SituacaoTributaria = 1 Then
                                    MovimentoCupomFiscalItem.CodigoAliquota = 2
                                'II
                                ElseIf rstMovCupomFiscalItem!SituacaoTributaria = 2 Then
                                    MovimentoCupomFiscalItem.CodigoAliquota = 1
                                'NN
                                ElseIf rstMovCupomFiscalItem!SituacaoTributaria = 3 Then
                                    MovimentoCupomFiscalItem.CodigoAliquota = 3
                                '17%
                                ElseIf rstMovCupomFiscalItem!SituacaoTributaria = 7 Then
                                    MovimentoCupomFiscalItem.CodigoAliquota = 5
                                Else
                                    MsgBox "Aliquota?"
                                End If
                                MovimentoCupomFiscalItem.CodigoGrupo = Produto.CodigoGrupo
                                If Not MovimentoCupomFiscalItem.Incluir Then
                                    MsgBox "Não foi possível incluir item do cupom fiscal.", vbInformation, "Erro de Integridade!"
                                End If
                            Else
                                MsgBox "Produto não localizado"
                            End If

                        
                            rstMovCupomFiscalItem.MoveNext
                        Loop
                    End If
                
                End If
                .MoveNext
            Loop
        End If
    End With
    
    
    
    
    
    
    
    cnnSGPDados.Close
End Sub

