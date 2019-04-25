VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMovimentoCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mData As Date
Private mNumeroMovimento As Long
Private mValor As Currency
Private mNumeroDocumento As String
Private mCodigoHistorico As Integer
Private mComplemento As String
Private mNumeroContaDebito As String
Private mNumeroContaCredito As String
Private mTipoMovimento As Integer
Private mFluxoCaixa As Boolean
Private mCodigoUsuario As Integer

Private rsMovCaixa As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get Empresa() As Integer
    Empresa = mEmpresa
End Property
Public Property Let Empresa(ByVal Valor As Integer)
    mEmpresa = Valor
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


Public Property Get NumeroContaDebito() As String
    NumeroContaDebito = mNumeroContaDebito
End Property
Public Property Let NumeroContaDebito(ByVal Valor As String)
    mNumeroContaDebito = Valor
End Property


Public Property Get NumeroContaCredito() As String
    NumeroContaCredito = mNumeroContaCredito
End Property
Public Property Let NumeroContaCredito(ByVal Valor As String)
    mNumeroContaCredito = Valor
End Property


Public Property Get TipoMovimento() As Integer
    TipoMovimento = mTipoMovimento
End Property
Public Property Let TipoMovimento(ByVal Valor As Integer)
    mTipoMovimento = Valor
End Property


Public Property Get FluxoCaixa() As Boolean
    FluxoCaixa = mFluxoCaixa
End Property
Public Property Let FluxoCaixa(ByVal Valor As Boolean)
    mFluxoCaixa = Valor
End Property


Public Property Get CodigoUsuario() As Integer
    CodigoUsuario = mCodigoUsuario
End Property
Public Property Let CodigoUsuario(ByVal Valor As Integer)
    mCodigoUsuario = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarRegistro(ByVal nEmpresa As Integer, ByVal nData As Date, ByVal nNumeroMovimento As Long) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarRegistro = False
    xCondicao = " WHERE Empresa = " & nEmpresa
    xCondicao = xCondicao & " AND Data = " & preparaData(nData)
    xCondicao = xCondicao & " AND [Numero do Movimento] = " & nNumeroMovimento
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarRegistro = Localizar(1)
    If LocalizarRegistro Then
        AtribuiValor
    End If
    rsMovCaixa.Close
    Set rsMovCaixa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarRegistroEspecial(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pNumeroConta As String, ByVal pDebitoCredito As String) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarRegistroEspecial = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND Data = " & preparaData(pData)
    If pDebitoCredito = "D" Then
        xCondicao = xCondicao & " AND [Numero da Conta Debito] = " & preparaTexto(pNumeroConta)
    Else
        xCondicao = xCondicao & " AND [Numero da Conta Credito] = " & preparaTexto(pNumeroConta)
    End If
    xCondicao = xCondicao & " AND [Tipo do Movimento] = " & 2
    gSQL = PreparaSQL(xCondicao, "")
    If Localizar(1) Then
        LocalizarRegistroEspecial = True
        AtribuiValor
    End If
    rsMovCaixa.Close
    Set rsMovCaixa = Nothing
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
    xSQL = xSQL & " AND Data <= " & preparaData(mData)
    'xSQL = xSQL & " AND [Numero do Movimento] <= " & mNumeroMovimento
    xOrdem = "ORDER BY Data DESC, [Numero do Movimento] DESC"
    gSQL = PreparaSQL(xSQL, xOrdem)
    If Localizar(99) Then
        If rsMovCaixa.RecordCount > 1 Then
            Do Until rsMovCaixa.EOF
                If xRegIgual = False Then
                    If rsMovCaixa("Data").Value = mData And rsMovCaixa("Numero do Movimento").Value = mNumeroMovimento Then
                        xRegIgual = True
                    End If
                Else
                    LocalizarAnterior = True
                    Exit Do
                End If
                rsMovCaixa.MoveNext
            Loop
        End If
    End If
    If LocalizarAnterior Then
        AtribuiValor
    End If
    rsMovCaixa.Close
    Set rsMovCaixa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro(ByVal nEmpresa As Integer) As Boolean
Dim xCondicao As String
Dim xSQL As String
Dim xOrdem As String

On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xSQL = " WHERE Empresa = " & nEmpresa
    xOrdem = "ORDER BY Data ASC, [Numero do Movimento] ASC"
    gSQL = PreparaSQL(xSQL, xOrdem)
    LocalizarPrimeiro = Localizar(1)
    If LocalizarPrimeiro Then
        AtribuiValor
    End If
    rsMovCaixa.Close
    Set rsMovCaixa = Nothing
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
    xSQL = xSQL & " AND Data >= " & preparaData(mData)
    'xSQL = xSQL & " AND [Numero do Movimento] >= " & mNumeroMovimento
    xOrdem = "ORDER BY Data ASC, [Numero do Movimento] ASC"
    gSQL = PreparaSQL(xSQL, xOrdem)
    If Localizar(99) Then
        If rsMovCaixa.RecordCount > 1 Then
            Do Until rsMovCaixa.EOF
                If xRegIgual = False Then
                    If rsMovCaixa("Data").Value = mData And rsMovCaixa("Numero do Movimento").Value = mNumeroMovimento Then
                        xRegIgual = True
                    End If
                Else
                    LocalizarProximo = True
                    Exit Do
                End If
                rsMovCaixa.MoveNext
            Loop
        End If
    End If
    If LocalizarProximo Then
        AtribuiValor
    End If
    rsMovCaixa.Close
    Set rsMovCaixa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal nEmpresa As Integer) As Boolean
Dim xCondicao As String
Dim xSQL As String
Dim xOrdem As String

On Error GoTo trata_erro

    LocalizarUltimo = False
    xSQL = " WHERE Empresa = " & nEmpresa
    xOrdem = "ORDER BY Data DESC, [Numero do Movimento] DESC"
    gSQL = PreparaSQL(xSQL, xOrdem)
    LocalizarUltimo = Localizar(1)
    If LocalizarUltimo Then
        AtribuiValor
    End If
    rsMovCaixa.Close
    Set rsMovCaixa = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Long
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = 0
    mNumeroMovimento = ProximoRegistro(mEmpresa)
    gSQL = "INSERT INTO Movimento_Caixa ( Empresa, Data, [Numero do Movimento], Valor, "
    gSQL = gSQL & "[Numero do Documento], [Codigo do Historico], Complemento, "
    gSQL = gSQL & "[Numero da Conta Debito], [Numero da Conta Credito], [Tipo do Movimento], "
    gSQL = gSQL & "[Fluxo de Caixa], [Codigo do Usuario] ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlData(1, mData, ", ")
    Call sqlNumero(1, mNumeroMovimento, ", ")
    Call sqlValor(1, mValor, ", ")
    Call sqlTexto(1, mNumeroDocumento, ", ")
    Call sqlNumero(1, mCodigoHistorico, ", ")
    Call sqlTexto(1, mComplemento, ", ")
    Call sqlTexto(1, mNumeroContaDebito, ", ")
    Call sqlTexto(1, mNumeroContaCredito, ", ")
    Call sqlNumero(1, mTipoMovimento, ", ")
    Call sqlBoolean(1, mFluxoCaixa, ", ")
    Call sqlNumero(1, mCodigoUsuario, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        If LoopIncluiSaldo Then
            Incluir = mNumeroMovimento
        Else
            MsgBox "Erro na atualização do saldo."
        End If
    End If
    Exit Function

Err_Incluir:
    MsgBox Error
End Function
Private Function LoopIncluiSaldo() As Boolean
    LoopIncluiSaldo = False
    If mNumeroContaDebito <> "" Then
        LoopIncluiSaldo = IncluiSaldo(mNumeroContaDebito, "D")
    End If
    If mNumeroContaCredito <> "" Then
        LoopIncluiSaldo = IncluiSaldo(mNumeroContaCredito, "C")
    End If
End Function

Private Function IncluiSaldo(ByVal pNumeroConta As String, ByVal pDebitoCredito As String) As Boolean
Dim lRecordsAffected As Long
Dim xSaldo As Currency
Dim xValor As Currency
Dim xOperador As String
Dim xGrupoConta As Integer

On Error GoTo Err_IncluiSaldo

    IncluiSaldo = False
    
    xGrupoConta = Val(Mid(pNumeroConta, 1, 1))
    If xGrupoConta = 1 Or xGrupoConta = 4 Then
        If pDebitoCredito = "C" Then
            xValor = -mValor
        Else
            xValor = mValor
        End If
    ElseIf xGrupoConta = 2 Or xGrupoConta = 3 Then
        If pDebitoCredito = "C" Then
            xValor = mValor
        Else
            xValor = -mValor
        End If
    End If

    'Busca Saldo Anterior a Data Atual
    xSaldo = 0
    gSQL = ""
    gSQL = gSQL & "   SELECT Saldo"
    gSQL = gSQL & "     FROM Saldo_Conta"
    gSQL = gSQL & "    WHERE Empresa = " & mEmpresa
    gSQL = gSQL & "      AND [Codigo da Conta] = " & preparaTexto(pNumeroConta)
    gSQL = gSQL & "      AND Data < " & preparaData(mData)
    gSQL = gSQL & " ORDER BY Data DESC"
    If Localizar(1) Then
        xSaldo = rsMovCaixa("Saldo").Value
    End If
    rsMovCaixa.Close
    Set rsMovCaixa = Nothing

    'Busca Saldo na Data Atual
    gSQL = ""
    gSQL = gSQL & "   SELECT Saldo"
    gSQL = gSQL & "     FROM Saldo_Conta"
    gSQL = gSQL & "    WHERE Empresa = " & mEmpresa
    gSQL = gSQL & "      AND [Codigo da Conta] = " & preparaTexto(pNumeroConta)
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
        Call sqlTexto(1, pNumeroConta, ", ")
        Call sqlData(1, mData, ", ")
        Call sqlValor(1, xSaldo, " )")
        gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
        If lRecordsAffected = 0 Then
            Exit Function
        End If
    End If
    rsMovCaixa.Close
    Set rsMovCaixa = Nothing

    'Atualiza Saldo com Data >= ou = a Atual
    'Conforme variavel xOperador
    gSQL = "UPDATE Saldo_Conta SET "
    Call sqlValor(2, "Saldo = Saldo + ", xValor)
    gSQL = gSQL & " WHERE Empresa = " & mEmpresa
    gSQL = gSQL & "   AND [Codigo da Conta] = " & preparaTexto(pNumeroConta)
    gSQL = gSQL & "   AND Data " & xOperador & " " & preparaData(mData)
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    
    IncluiSaldo = True
    Exit Function

Err_IncluiSaldo:
    MsgBox "Erro na Rotina de Saldo", vbInformation, "IncluiSaldo"
    Exit Function
End Function


Public Function Alterar(ByVal nEmpresa As Integer, ByVal nData As Date, ByVal nNumeroMovimento As Long) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    If Not LoopExcluiSaldo(nEmpresa, nData, nNumeroMovimento) Then
        Exit Function
    End If
    gSQL = "UPDATE Movimento_Caixa SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlData(2, ", Data = ", mData)
    Call sqlNumero(2, ", [Numero do Movimento] = ", mNumeroMovimento)
    Call sqlValor(2, ", Valor = ", mValor)
    Call sqlTexto(2, ", [Numero do Documento] = ", mNumeroDocumento)
    Call sqlNumero(2, ", [Codigo do Historico] = ", mCodigoHistorico)
    Call sqlTexto(2, ", Complemento = ", mComplemento)
    Call sqlTexto(2, ", [Numero da Conta Debito] = ", mNumeroContaDebito)
    Call sqlTexto(2, ", [Numero da Conta Credito] = ", mNumeroContaCredito)
    Call sqlNumero(2, ", [Tipo do Movimento] = ", mTipoMovimento)
    Call sqlBoolean(2, ", [Fluxo de Caixa] = ", mFluxoCaixa)
    Call sqlNumero(2, ", [Codigo do Usuario] = ", mCodigoUsuario)
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(nData)
    gSQL = gSQL & " AND [Numero do Movimento] = " & nNumeroMovimento
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        If LoopIncluiSaldo Then
            Alterar = True
        Else
            MsgBox "Erro na atualização do saldo."
        End If
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal nEmpresa As Integer, ByVal nData As Date, ByVal nNumeroMovimento As Long) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    If Not LoopExcluiSaldo(nEmpresa, nData, nNumeroMovimento) Then
        Exit Function
    End If
    gSQL = "DELETE FROM Movimento_Caixa"
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(nData)
    gSQL = gSQL & " AND [Numero do Movimento] = " & nNumeroMovimento
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


Private Function LoopExcluiSaldo(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pNumeroMovimento As Long) As Boolean
Dim xContaDebito As String
Dim xContaCredito As String
Dim xValor As Currency

On Error GoTo Err_LoopExcluiSaldo
    LoopExcluiSaldo = False
    xContaDebito = ""
    xContaCredito = ""
    xValor = 0
    
    'Busca registro atual do Movimento_Caixa
    gSQL = ""
    gSQL = gSQL & "   SELECT Data, Valor, [Numero da Conta Debito], [Numero da Conta Credito]"
    gSQL = gSQL & "     FROM Movimento_Caixa"
    gSQL = gSQL & "    WHERE Empresa = " & pEmpresa
    gSQL = gSQL & "      AND Data = " & preparaData(pData)
    gSQL = gSQL & "      AND [Numero do Movimento] = " & pNumeroMovimento
    If Localizar(1) Then
        xContaDebito = rsMovCaixa("Numero da Conta Debito").Value
        xContaCredito = rsMovCaixa("Numero da Conta Credito").Value
        xValor = rsMovCaixa("Valor").Value
    End If
    rsMovCaixa.Close
    Set rsMovCaixa = Nothing
    
    If xContaDebito <> "" Then
        LoopExcluiSaldo = ExcluiSaldo(pEmpresa, pData, xContaDebito, "D", xValor)
    End If
    If xContaCredito <> "" Then
        LoopExcluiSaldo = ExcluiSaldo(pEmpresa, pData, xContaCredito, "C", xValor)
    End If
    Exit Function

Err_LoopExcluiSaldo:
    MsgBox "Erro na Rotina de Saldo", vbInformation, "LoopExcluiSaldo"
    Exit Function
End Function


Public Function RecalculaSaldo(ByVal pEmpresa As Integer) As Boolean
Dim rsMovimento As New adodb.Recordset


On Error GoTo Erro_Rotina

    RecalculaSaldo = False

    Call RecalculaSaldo_ExcluiSaldo(Empresa, "D")
    Call RecalculaSaldo_ExcluiSaldo(Empresa, "C")
    
    'Busca Movimento_Caixa da empresa
    gSQL = ""
    gSQL = gSQL & "SELECT Data, [Numero do Movimento]"
    gSQL = gSQL & "  FROM Movimento_Caixa"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & "   AND Data >= " & preparaData("01/12/2004")
    gSQL = gSQL & " ORDER BY Data, [Numero do Movimento]"
    Set rsMovimento = New adodb.Recordset
    rsMovimento.CursorLocation = adUseClient
    rsMovimento.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovimento.EOF Then
        Do Until rsMovimento.EOF
            If LocalizarRegistro(pEmpresa, rsMovimento("Data").Value, rsMovimento("Numero do Movimento").Value) Then
                Call LoopIncluiSaldo
            End If
            rsMovimento.MoveNext
        Loop
    End If
    rsMovimento.Close
    Set rsMovimento = Nothing

    RecalculaSaldo = True
    Exit Function

Erro_Rotina:
    MsgBox "Erro interno na Rotina!", vbInformation, "RecalculaSaldo"
    Exit Function
End Function


Private Function RecalculaSaldo_ExcluiSaldo(ByVal pEmpresa As Integer, ByVal pDebitoCredito As String) As Boolean
Dim lRecordsAffected As Long
Dim xString As String
Dim rsMovimento As New adodb.Recordset


On Error GoTo Erro_Rotina

    RecalculaSaldo_ExcluiSaldo = False

    If pDebitoCredito = "D" Then
        xString = "[Numero da Conta Debito]"
    Else
        xString = "[Numero da Conta Credito]"
    End If
    
    'Busca Movimento_Caixa da empresa
    gSQL = ""
    gSQL = gSQL & "SELECT " & xString & " AS [Numero da Conta]"
    gSQL = gSQL & "  FROM Movimento_Caixa"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & "   AND " & xString & " <> " & preparaTexto("")
    gSQL = gSQL & " GROUP BY " & xString
    Set rsMovimento = New adodb.Recordset
    rsMovimento.CursorLocation = adUseClient
    rsMovimento.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovimento.EOF Then
        Do Until rsMovimento.EOF
            'Esclui os Saldos da conta específica
            gSQL = "DELETE * FROM Saldo_Conta "
            gSQL = gSQL & " WHERE Empresa = " & pEmpresa
            gSQL = gSQL & "   AND [Codigo da Conta] = " & preparaTexto(rsMovimento("Numero da Conta").Value)
            gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
            rsMovimento.MoveNext
        Loop
    End If
    rsMovimento.Close
    Set rsMovimento = Nothing
    Exit Function

Erro_Rotina:
    MsgBox "Erro interno na Rotina!", vbInformation, "RecalculaSaldo_ExcluiSaldo"
    Exit Function
End Function


Public Function ExcluiSaldo(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pNumeroConta As String, ByVal pDebitoCredito As String, ByVal pValor As Currency) As Boolean
Dim lRecordsAffected As Long
Dim xGrupoConta As Integer

On Error GoTo Err_ExcluiSaldo

    ExcluiSaldo = False

    xGrupoConta = Val(Mid(pNumeroConta, 1, 1))
    If xGrupoConta = 1 Or xGrupoConta = 4 Then
        If pDebitoCredito = "C" Then
            pValor = -pValor
        Else
            pValor = pValor
        End If
    ElseIf xGrupoConta = 2 Or xGrupoConta = 3 Then
        If pDebitoCredito = "C" Then
            pValor = pValor
        Else
            pValor = -pValor
        End If
    End If
    

    'Atualiza Saldo com Data >= a Atual
    gSQL = "UPDATE Saldo_Conta SET "
    Call sqlValor(2, "Saldo = Saldo - ", pValor)
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & "   AND [Codigo da Conta] = " & preparaTexto(pNumeroConta)
    gSQL = gSQL & "   AND Data >= " & preparaData(pData)
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
    PreparaSQL = Trim("SELECT Empresa, Data, [Numero do Movimento], Valor, [Numero do Documento], [Codigo do Historico], Complemento, [Numero da Conta Debito], [Numero da Conta Credito], [Tipo do Movimento], [Fluxo de Caixa], [Codigo do Usuario] FROM Movimento_Caixa" & " " & xCondicao & " " & xTipoOrdem)
End Function


Public Function ProximoRegistro(ByVal nEmpresa As Integer) As Long

On Error GoTo Err_ProximoRegistro

    'Pega número do último movimento
    ProximoRegistro = 1
    gSQL = ""
    gSQL = gSQL & "SELECT [Numero do Movimento]"
    gSQL = gSQL & "  FROM Movimento_Caixa"
    gSQL = gSQL & " WHERE Empresa = " & mEmpresa
    gSQL = gSQL & "   AND Data = " & preparaData(mData)
    gSQL = gSQL & " ORDER BY [Numero do Movimento] DESC"
    If Localizar(1) Then
        ProximoRegistro = rsMovCaixa("Numero do Movimento").Value + 1
    End If
    rsMovCaixa.Close
    Set rsMovCaixa = Nothing
    Exit Function

Err_ProximoRegistro:
End Function




Public Function TotalData(ByVal pEmpresa As Integer, ByVal pDataInicial As Date, ByVal pDataFinal As Date, ByVal pDebitoCredito As String) As Currency

On Error GoTo Err_TotalData
    TotalData = 0
    
    gSQL = ""
    gSQL = gSQL & "SELECT SUM(Valor) AS Total" ', [Numero da Conta Debito], [Numero da Conta Credito]"
    gSQL = gSQL & "  FROM Movimento_Caixa"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & "   AND Data >= " & preparaData(pDataInicial)
    gSQL = gSQL & "   AND Data <= " & preparaData(pDataFinal)
    If pDebitoCredito = "C" Then
        gSQL = gSQL & "   AND [Numero da Conta Credito] <> " & preparaTexto("")
    Else
        gSQL = gSQL & "   AND [Numero da Conta Debito] <> " & preparaTexto("")
    End If
    If Localizar(0) Then
        If Not IsNull(rsMovCaixa("Total").Value) Then
            TotalData = rsMovCaixa("Total").Value
        End If
    End If
    rsMovCaixa.Close
    Set rsMovCaixa = Nothing
    Exit Function

Err_TotalData:
    MsgBox "Erro na Rotina: TotalData", vbInformation
    Exit Function
End Function



Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsMovCaixa = New adodb.Recordset
    rsMovCaixa.CursorLocation = adUseClient
    gSQL = Trim(gSQL)
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsMovCaixa.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovCaixa.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsMovCaixa("Empresa").Value
    Data = rsMovCaixa("Data").Value
    NumeroMovimento = rsMovCaixa("Numero do Movimento").Value
    Valor = rsMovCaixa("Valor").Value
    NumeroDocumento = rsMovCaixa("Numero do Documento").Value
    CodigoHistorico = rsMovCaixa("Codigo do Historico").Value
    Complemento = rsMovCaixa("Complemento").Value
    NumeroContaDebito = rsMovCaixa("Numero da Conta Debito").Value
    NumeroContaCredito = rsMovCaixa("Numero da Conta Credito").Value
    TipoMovimento = rsMovCaixa("Tipo do Movimento").Value
    FluxoCaixa = rsMovCaixa("Fluxo de Caixa").Value
    CodigoUsuario = rsMovCaixa("Codigo do Usuario").Value
End Sub


Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function


