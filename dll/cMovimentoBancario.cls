VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMovimentoBancario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mCodigoPortador As String
Private mData As Date
Private mNumeroMovimento As Long
Private mValor As Currency
Private mDebitoouCredito As String
Private mNumeroDocumento As String
Private mCodigoHistorico As Integer
Private mComplemento As String

Private rsMovBancario As New adodb.Recordset
Private rsSaldo As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get Empresa() As Integer
    Empresa = mEmpresa
End Property
Public Property Let Empresa(ByVal Valor As Integer)
    mEmpresa = Valor
End Property


Public Property Get CodigoPortador() As String
    CodigoPortador = mCodigoPortador
End Property
Public Property Let CodigoPortador(ByVal Valor As String)
    mCodigoPortador = Valor
End Property


Public Property Get Data() As Date
    Data = mData
End Property
Public Property Let Data(ByVal Valor As Date)
    mData = Valor
End Property


Public Property Get NumeroMovimento() As Long
    NumeroMovimento = mNumeroMovimento
End Property
Public Property Let NumeroMovimento(ByVal Valor As Long)
    mNumeroMovimento = Valor
End Property


Public Property Get Valor() As Currency
    Valor = mValor
End Property
Public Property Let Valor(ByVal Valor As Currency)
    mValor = Valor
End Property


Public Property Get DebitoouCredito() As String
    DebitoouCredito = mDebitoouCredito
End Property
Public Property Let DebitoouCredito(ByVal Valor As String)
    mDebitoouCredito = Valor
End Property


Public Property Get NumeroDocumento() As String
    NumeroDocumento = mNumeroDocumento
End Property
Public Property Let NumeroDocumento(ByVal Valor As String)
    mNumeroDocumento = Valor
End Property


Public Property Get CodigoHistorico() As Integer
    CodigoHistorico = mCodigoHistorico
End Property
Public Property Let CodigoHistorico(ByVal Valor As Integer)
    mCodigoHistorico = Valor
End Property


Public Property Get Complemento() As String
    Complemento = mComplemento
End Property
Public Property Let Complemento(ByVal Valor As String)
    mComplemento = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarRegistro(ByVal nEmpresa As Integer, ByVal nCodigoPortador As String, ByVal nData As Date, ByVal nNumeroMovimento As Long) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarRegistro = False
    xCondicao = " WHERE Empresa = " & nEmpresa
    xCondicao = xCondicao & " AND [Codigo do Portador] = " & preparaTexto(nCodigoPortador)
    xCondicao = xCondicao & " AND Data = " & preparaData(nData)
    xCondicao = xCondicao & " AND [Numero do Movimento] = " & nNumeroMovimento
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarRegistro = Localizar(1)
    If LocalizarRegistro Then
        AtribuiValor
    End If
    rsMovBancario.Close
    Set rsMovBancario = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarAnterior() As Boolean
Dim xSQL As String
Dim xOrdem As String
Dim xRegIgual As Boolean

On Error GoTo trata_erro

    LocalizarAnterior = False
    xRegIgual = False
    xSQL = " WHERE Empresa = " & mEmpresa
    xSQL = xSQL & " AND [Codigo do Portador] = " & preparaTexto(mCodigoPortador)
    xSQL = xSQL & " AND Data <= " & preparaData(mData)
    'xSQL = xSQL & " AND [Numero do Movimento] <= " & mNumeroMovimento
    xOrdem = "ORDER BY Data DESC, [Numero do Movimento] DESC"
    gSQL = PreparaSQL(xSQL, xOrdem)
    If Localizar(99) Then
        If rsMovBancario.RecordCount > 1 Then
            Do Until rsMovBancario.EOF
                If xRegIgual = False Then
                    If rsMovBancario("Data").Value = mData And rsMovBancario("Numero do Movimento").Value = mNumeroMovimento Then
                        xRegIgual = True
                    End If
                Else
                    LocalizarAnterior = True
                    Exit Do
                End If
                rsMovBancario.MoveNext
            Loop
        End If
    End If
    If LocalizarAnterior Then
        AtribuiValor
    End If
    rsMovBancario.Close
    Set rsMovBancario = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarNome(ByVal nNome As String) As Boolean

On Error GoTo trata_erro
    
    LocalizarNome = False
    gSQL = PreparaSQL(" WHERE Nome = " & preparaTexto(nNome), "")
    LocalizarNome = Localizar(1)
    If LocalizarNome Then
        AtribuiValor
    End If
    rsMovBancario.Close
    Set rsMovBancario = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro(ByVal nEmpresa As Integer, ByVal nCodigoPortador As String) As Boolean
Dim xCondicao As String
Dim xSQL As String
Dim xOrdem As String

On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xSQL = " WHERE Empresa = " & nEmpresa
    xSQL = xSQL & " AND [Codigo do Portador] = " & preparaTexto(nCodigoPortador)
    xOrdem = "ORDER BY Data ASC, [Numero do Movimento] ASC"
    gSQL = PreparaSQL(xSQL, xOrdem)
    LocalizarPrimeiro = Localizar(1)
    If LocalizarPrimeiro Then
        AtribuiValor
    End If
    rsMovBancario.Close
    Set rsMovBancario = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarProximo() As Boolean
Dim xSQL As String
Dim xOrdem As String
Dim xRegIgual As Boolean

On Error GoTo trata_erro

    LocalizarProximo = False
    xRegIgual = False
    xSQL = " WHERE Empresa = " & mEmpresa
    xSQL = xSQL & " AND [Codigo do Portador] = " & preparaTexto(mCodigoPortador)
    xSQL = xSQL & " AND Data >= " & preparaData(mData)
    'xSQL = xSQL & " AND [Numero do Movimento] >= " & mNumeroMovimento
    xOrdem = "ORDER BY Data ASC, [Numero do Movimento] ASC"
    gSQL = PreparaSQL(xSQL, xOrdem)
    If Localizar(99) Then
        If rsMovBancario.RecordCount > 1 Then
            Do Until rsMovBancario.EOF
                If xRegIgual = False Then
                    If rsMovBancario("Data").Value = mData And rsMovBancario("Numero do Movimento").Value = mNumeroMovimento Then
                        xRegIgual = True
                    End If
                Else
                    LocalizarProximo = True
                    Exit Do
                End If
                rsMovBancario.MoveNext
            Loop
        End If
    End If
    If LocalizarProximo Then
        AtribuiValor
    End If
    rsMovBancario.Close
    Set rsMovBancario = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimoGeral(ByVal nEmpresa As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimoGeral = False
    xCondicao = " WHERE Empresa = " & nEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa DESC, [Codigo do Portador] DESC, [Numero do Movimento] DESC")
    LocalizarUltimoGeral = Localizar(1)
    If LocalizarUltimoGeral Then
        AtribuiValor
    End If
    rsMovBancario.Close
    Set rsMovBancario = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal nEmpresa As Integer, ByVal nCodigoPortador As String) As Boolean
Dim xCondicao As String
Dim xSQL As String
Dim xOrdem As String

On Error GoTo trata_erro

    LocalizarUltimo = False
    xSQL = " WHERE Empresa = " & nEmpresa
    xSQL = xSQL & " AND [Codigo do Portador] = " & preparaTexto(nCodigoPortador)
    xOrdem = "ORDER BY Data DESC, [Numero do Movimento] DESC"
    gSQL = PreparaSQL(xSQL, xOrdem)
    LocalizarUltimo = Localizar(1)
    If LocalizarUltimo Then
        AtribuiValor
    End If
    rsMovBancario.Close
    Set rsMovBancario = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Long
Dim lRecordsAffected As Long
Dim xNumeroMovimento As Long

On Error GoTo Err_Incluir

    'Pega número do último movimento
    NumeroMovimento = 1
    gSQL = ""
    gSQL = gSQL & "   SELECT [Numero do Movimento]"
    gSQL = gSQL & "     FROM MovimentoBancario"
    gSQL = gSQL & "    WHERE Empresa = " & mEmpresa
    gSQL = gSQL & "      AND [Codigo do Portador] = " & preparaTexto(mCodigoPortador)
    gSQL = gSQL & "      AND Data = " & preparaData(mData)
    gSQL = gSQL & " ORDER BY [Numero do Movimento] DESC"
    If Localizar(1) Then
        NumeroMovimento = rsMovBancario("Numero do Movimento").Value + 1
    End If
    rsMovBancario.Close
    Set rsMovBancario = Nothing


    Incluir = 0
    gSQL = "INSERT INTO MovimentoBancario ( Empresa, [Codigo do Portador], Data, [Numero do Movimento], Valor, "
    gSQL = gSQL & "[Debito ou Credito], [Numero do Documento], [Codigo do Historico], Complemento ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlTexto(1, mCodigoPortador, ", ")
    Call sqlData(1, mData, ", ")
    Call sqlNumero(1, mNumeroMovimento, ", ")
    Call sqlValor(1, mValor, ", ")
    Call sqlTexto(1, mDebitoouCredito, ", ")
    Call sqlTexto(1, mNumeroDocumento, ", ")
    Call sqlNumero(1, mCodigoHistorico, ", ")
    Call sqlTexto(1, mComplemento, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        If IncluiSaldo Then
            Incluir = mNumeroMovimento
        Else
            MsgBox "Erro na atualização do saldo."
        End If
    End If
    Exit Function

Err_Incluir:
End Function


Public Function BuscaSaldoAnterior(ByVal nEmpresa As Integer, ByVal nCodigoPortador As String, ByVal nData As Date) As Currency

On Error GoTo ErroInterno

    BuscaSaldoAnterior = 0
    'Busca Saldo Anterior a Data Atual
    gSQL = ""
    gSQL = gSQL & "   SELECT Saldo"
    gSQL = gSQL & "     FROM Saldo_Conta"
    gSQL = gSQL & "    WHERE Empresa = " & nEmpresa
    gSQL = gSQL & "      AND [Codigo da Conta] = " & preparaTexto(nCodigoPortador)
    gSQL = gSQL & "      AND Data < " & preparaData(nData)
    gSQL = gSQL & " ORDER BY Data DESC"
    If Localizar(1) Then
        BuscaSaldoAnterior = rsMovBancario("Saldo").Value
    End If
    rsMovBancario.Close
    Set rsMovBancario = Nothing
    Exit Function

ErroInterno:
    MsgBox "Erro ao buscar o saldo anterior", vbInformation, "BuscaSaldoAnterior"
    Exit Function
End Function


Public Function IncluiSaldo() As Boolean
Dim lRecordsAffected As Long
Dim xSaldo As Currency
Dim xValor As Currency
Dim xOperador As String

On Error GoTo Err_IncluiSaldo

    IncluiSaldo = False
    If mDebitoouCredito = "C" Then
        xValor = mValor
    Else
        xValor = -mValor
    End If

    'Busca Saldo Anterior a Data Atual
    xSaldo = 0
    gSQL = ""
    gSQL = gSQL & "   SELECT Saldo"
    gSQL = gSQL & "     FROM Saldo_Conta"
    gSQL = gSQL & "    WHERE Empresa = " & mEmpresa
    gSQL = gSQL & "      AND [Codigo da Conta] = " & preparaTexto(mCodigoPortador)
    gSQL = gSQL & "      AND Data < " & preparaData(mData)
    gSQL = gSQL & " ORDER BY Data DESC"
    If Localizar(1) Then
        xSaldo = rsMovBancario("Saldo").Value
    End If
    rsMovBancario.Close
    Set rsMovBancario = Nothing

    'Busca Saldo na Data Atual
    gSQL = ""
    gSQL = gSQL & "   SELECT Saldo"
    gSQL = gSQL & "     FROM Saldo_Conta"
    gSQL = gSQL & "    WHERE Empresa = " & mEmpresa
    gSQL = gSQL & "      AND [Codigo da Conta] = " & preparaTexto(mCodigoPortador)
    gSQL = gSQL & "      AND Data = " & preparaData(mData)
    gSQL = gSQL & " ORDER BY Data DESC"
    If Localizar(1) Then
        xOperador = ">="
    Else
        'Inclui Saldo na Data Atual
        xOperador = ">"
        xSaldo = xSaldo + xValor
        gSQL = "INSERT INTO Saldo_Conta ( Empresa, [Codigo da Conta], Data, Saldo ) VALUES ( "
        Call sqlNumero(1, mEmpresa, ", ")
        Call sqlTexto(1, mCodigoPortador, ", ")
        Call sqlData(1, mData, ", ")
        Call sqlValor(1, xSaldo, " )")
        gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
        If lRecordsAffected = 0 Then
            Exit Function
        End If
    End If
    rsMovBancario.Close
    Set rsMovBancario = Nothing

    'Atualiza Saldo com Data >= ou = a Atual
    'Conforme variavel xOperador
    gSQL = "UPDATE Saldo_Conta SET "
    Call sqlValor(2, "Saldo = Saldo + ", xValor)
    gSQL = gSQL & " WHERE Empresa = " & mEmpresa
    gSQL = gSQL & " AND [Codigo da Conta] = " & preparaTexto(mCodigoPortador)
    gSQL = gSQL & " AND Data " & xOperador & " " & preparaData(mData)
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    
    IncluiSaldo = True
    Exit Function

Err_IncluiSaldo:
    MsgBox "Erro na Rotina de Saldo", vbInformation, "IncluiSaldo"
    Exit Function
End Function


Public Function Alterar(ByVal nEmpresa As Integer, ByVal nCodigoPortador As String, ByVal nData As Date, ByVal nNumeroMovimento As Long) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    If Not ExcluiSaldo(nEmpresa, nCodigoPortador, nData, nNumeroMovimento) Then
        Exit Function
    End If
    gSQL = "UPDATE MovimentoBancario SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlTexto(2, ", [Codigo do Portador] = ", mCodigoPortador)
    Call sqlData(2, ", Data = ", mData)
    Call sqlNumero(2, ", [Numero do Movimento] = ", mNumeroMovimento)
    Call sqlValor(2, ", Valor = ", mValor)
    Call sqlTexto(2, ", [Debito ou Credito] = ", mDebitoouCredito)
    Call sqlTexto(2, ", [Numero do Documento] = ", mNumeroDocumento)
    Call sqlNumero(2, ", [Codigo do Historico] = ", mCodigoHistorico)
    Call sqlTexto(2, ", Complemento = ", mComplemento)
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND [Codigo do Portador] = " & preparaTexto(nCodigoPortador)
    gSQL = gSQL & " AND Data = " & preparaData(nData)
    gSQL = gSQL & " AND [Numero do Movimento] = " & nNumeroMovimento
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        If IncluiSaldo Then
            Alterar = True
        Else
            MsgBox "Erro na atualização do saldo."
        End If
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal nEmpresa As Integer, ByVal nCodigoPortador As String, ByVal nData As Date, ByVal nNumeroMovimento As Long) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    If Not ExcluiSaldo(nEmpresa, nCodigoPortador, nData, nNumeroMovimento) Then
        Exit Function
    End If
    gSQL = "DELETE FROM MovimentoBancario"
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND [Codigo do Portador] = " & preparaTexto(nCodigoPortador)
    gSQL = gSQL & " AND Data = " & preparaData(nData)
    gSQL = gSQL & " AND [Numero do Movimento] = " & nNumeroMovimento
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


Public Function ExcluiSaldo(ByVal nEmpresa As Integer, ByVal nCodigoPortador As String, ByVal nData As Date, ByVal nNumeroMovimento As Long) As Boolean
Dim lRecordsAffected As Long
Dim xValor As Currency

On Error GoTo Err_ExcluiSaldo

    ExcluiSaldo = False

    'Busca registro atual do MovimentoBancario
    gSQL = ""
    gSQL = gSQL & "   SELECT Data, Valor, [Debito ou Credito]"
    gSQL = gSQL & "     FROM MovimentoBancario"
    gSQL = gSQL & "    WHERE Empresa = " & nEmpresa
    gSQL = gSQL & "      AND [Codigo do Portador] = " & preparaTexto(nCodigoPortador)
    gSQL = gSQL & "      AND Data = " & preparaData(nData)
    gSQL = gSQL & "      AND [Numero do Movimento] = " & nNumeroMovimento
    If Localizar(1) Then
        If rsMovBancario("Debito ou Credito").Value = "C" Then
            xValor = rsMovBancario("Valor").Value
        Else
            xValor = -rsMovBancario("Valor").Value
        End If
    End If
    rsMovBancario.Close
    Set rsMovBancario = Nothing

    'Atualiza Saldo com Data >= a Atual
    gSQL = "UPDATE Saldo_Conta SET "
    Call sqlValor(2, "Saldo = Saldo - ", xValor)
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND [Codigo da Conta] = " & preparaTexto(nCodigoPortador)
    gSQL = gSQL & " AND Data >= " & preparaData(nData)
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected = 0 Then
        Exit Function
    End If
    
    ExcluiSaldo = True
    Exit Function

Err_ExcluiSaldo:
    MsgBox "Erro na Rotina de Saldo", vbInformation, "ExcluiSaldo"
    Exit Function
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, [Codigo do Portador], Data, [Numero do Movimento], Valor, [Debito ou Credito], [Numero do Documento], [Codigo do Historico], Complemento FROM MovimentoBancario" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsMovBancario = New adodb.Recordset
    rsMovBancario.CursorLocation = adUseClient
    gSQL = Trim(gSQL)
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsMovBancario.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovBancario.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Function AtualizaRecordsetSaldo(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordsetSaldo = False
    Set rsSaldo = New adodb.Recordset
    rsSaldo.CursorLocation = adUseClient
    gSQL = Trim(gSQL)
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsSaldo.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsSaldo.EOF Then
        AtualizaRecordsetSaldo = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsMovBancario("Empresa").Value
    CodigoPortador = rsMovBancario("Codigo do Portador").Value
    Data = rsMovBancario("Data").Value
    NumeroMovimento = rsMovBancario("Numero do Movimento").Value
    Valor = rsMovBancario("Valor").Value
    DebitoouCredito = rsMovBancario("Debito ou Credito").Value
    NumeroDocumento = rsMovBancario("Numero do Documento").Value
    CodigoHistorico = rsMovBancario("Codigo do Historico").Value
    Complemento = rsMovBancario("Complemento").Value
End Sub


Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function


Private Function LocalizarSaldo(ByVal xQtdRegistro As Integer) As Boolean
    LocalizarSaldo = False
    If AtualizaRecordsetSaldo(xQtdRegistro) Then
        LocalizarSaldo = True
    End If
End Function


Public Function AlteraSaldoRecalcula(ByVal nEmpresa As Integer, ByVal nCodigoPortador As String, ByVal nData As Date, ByVal nValor As Currency) As Boolean
Dim lRecordsAffected As Long
Dim xSaldo As Currency

On Error GoTo ErroRotina

    AlteraSaldoRecalcula = False

    'Busca Saldo Anterior a Data Atual
    xSaldo = 0
    gSQL = ""
    gSQL = gSQL & "   SELECT Saldo"
    gSQL = gSQL & "     FROM Saldo_Conta"
    gSQL = gSQL & "    WHERE Empresa = " & nEmpresa
    gSQL = gSQL & "      AND [Codigo da Conta] = " & preparaTexto(nCodigoPortador)
    gSQL = gSQL & "      AND Data < " & preparaData(nData)
    gSQL = gSQL & " ORDER BY Data DESC"
    If LocalizarSaldo(1) Then
        xSaldo = rsSaldo("Saldo").Value
    End If
    rsSaldo.Close
    Set rsSaldo = Nothing

    'Busca Saldo na Data Atual
    gSQL = ""
    gSQL = gSQL & "   SELECT Saldo"
    gSQL = gSQL & "     FROM Saldo_Conta"
    gSQL = gSQL & "    WHERE Empresa = " & nEmpresa
    gSQL = gSQL & "      AND [Codigo da Conta] = " & preparaTexto(nCodigoPortador)
    gSQL = gSQL & "      AND Data = " & preparaData(nData)
    gSQL = gSQL & " ORDER BY Data DESC"
    If LocalizarSaldo(1) Then
        'Altera o Saldo na Data Atual
        gSQL = "UPDATE Saldo_Conta SET "
        Call sqlValor(2, "Saldo = Saldo + ", nValor)
        gSQL = gSQL & " WHERE Empresa = " & nEmpresa
        gSQL = gSQL & " AND [Codigo da Conta] = " & preparaTexto(nCodigoPortador)
        gSQL = gSQL & " AND Data = " & preparaData(nData)
        gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    Else
        'Inclui Saldo na Data Atual
        xSaldo = xSaldo + nValor
        gSQL = "INSERT INTO Saldo_Conta ( Empresa, [Codigo da Conta], Data, Saldo ) VALUES ( "
        Call sqlNumero(1, nEmpresa, ", ")
        Call sqlTexto(1, nCodigoPortador, ", ")
        Call sqlData(1, nData, ", ")
        Call sqlValor(1, xSaldo, " )")
        gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    End If
    rsSaldo.Close
    Set rsSaldo = Nothing
    AlteraSaldoRecalcula = True
    Exit Function

ErroRotina:
    MsgBox Error, vbInformation, "AlteraSaldoRecalcula"
    Exit Function
End Function


Public Function RecalculaSaldo(ByVal nEmpresa As Integer, ByVal nCodigoPortador As String) As Boolean
Dim lRecordsAffected As Long
Dim xSaldo As Currency
Dim xValor As Currency
Dim xCondicao As String

On Error GoTo ErroRotina

    RecalculaSaldo = False
    xSaldo = 0
    
    
    'Exclui Saldos da Conta
    gSQL = "DELETE FROM Saldo_Conta"
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND [Codigo da Conta] = " & preparaTexto(nCodigoPortador)
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    
    
    'Seleciona Movimento da Conta
    xCondicao = " WHERE Empresa = " & nEmpresa
    xCondicao = xCondicao & " AND [Codigo do Portador] = " & preparaTexto(nCodigoPortador)
    gSQL = PreparaSQL(xCondicao, "ORDER BY Data, [Numero do Movimento]")
    If Localizar(0) Then
        If rsMovBancario.RecordCount > 0 Then
            Do Until rsMovBancario.EOF
                If rsMovBancario("Debito ou Credito").Value = "C" Then
                    xValor = rsMovBancario("Valor").Value
                Else
                    xValor = -rsMovBancario("Valor").Value
                End If
                Call AlteraSaldoRecalcula(nEmpresa, nCodigoPortador, rsMovBancario("Data").Value, xValor)
                rsMovBancario.MoveNext
            Loop
        End If
    End If
    rsMovBancario.Close
    Set rsMovBancario = Nothing
    RecalculaSaldo = True
    Exit Function

ErroRotina:
    MsgBox Error, vbInformation, "RecalculaSaldo"
    Exit Function
End Function

