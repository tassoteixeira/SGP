VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMovimentoChequeAvista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mDataEmissao As Date
Private mPeriodo As String
Private mTipoMovimento As String
Private mValor As Currency
Private mOrdemDigitacao As Integer
Private mCodigoBarra1 As String
Private mCodigoBarra2 As String
Private mCodigoBarra3 As String

Private rsMovChequeAvista As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get Empresa() As Integer
    Empresa = mEmpresa
End Property
Public Property Let Empresa(ByVal Valor As Integer)
    mEmpresa = Valor
End Property


Public Property Get DataEmissao() As Date
    DataEmissao = mDataEmissao
End Property
Public Property Let DataEmissao(ByVal Valor As Date)
    mDataEmissao = Valor
End Property


Public Property Get Periodo() As String
    Periodo = mPeriodo
End Property
Public Property Let Periodo(ByVal Valor As String)
    mPeriodo = Valor
End Property


Public Property Get TipoMovimento() As String
    TipoMovimento = mTipoMovimento
End Property
Public Property Let TipoMovimento(ByVal Valor As String)
    mTipoMovimento = Valor
End Property


Public Property Get Valor() As Currency
    Valor = mValor
End Property
Public Property Let Valor(ByVal Valor As Currency)
    mValor = Valor
End Property


Public Property Get OrdemDigitacao() As Integer
    OrdemDigitacao = mOrdemDigitacao
End Property
Public Property Let OrdemDigitacao(ByVal Valor As Integer)
    mOrdemDigitacao = Valor
End Property


Public Property Get CodigoBarra1() As String
    CodigoBarra1 = mCodigoBarra1
End Property
Public Property Let CodigoBarra1(ByVal Valor As String)
    mCodigoBarra1 = Valor
End Property


Public Property Get CodigoBarra2() As String
    CodigoBarra2 = mCodigoBarra2
End Property
Public Property Let CodigoBarra2(ByVal Valor As String)
    mCodigoBarra2 = Valor
End Property


Public Property Get CodigoBarra3() As String
    CodigoBarra3 = mCodigoBarra3
End Property
Public Property Let CodigoBarra3(ByVal Valor As String)
    mCodigoBarra3 = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal nEmpresa As Integer, ByVal nDataEmissao As Date, ByVal nPeriodo As String, ByVal nTipoMovimento As String, ByVal nOrdemDigitacao As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & nEmpresa
    xCondicao = xCondicao & " AND [Data de Emissao] = " & preparaData(nDataEmissao)
    xCondicao = xCondicao & " AND Periodo = " & preparaTexto(nPeriodo)
    xCondicao = xCondicao & " AND [Tipo do Movimento] = " & preparaTexto(nTipoMovimento)
    xCondicao = xCondicao & " AND [Ordem da Digitacao] = " & nOrdemDigitacao
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo Then
        AtribuiValor
    End If
    rsMovChequeAvista.Close
    Set rsMovChequeAvista = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarAnterior() As Boolean
Dim xCondicao As String
Dim xRegIgual As Boolean

On Error GoTo trata_erro

    LocalizarAnterior = False
    xRegIgual = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND [Data de Emissao] <= " & preparaData(mDataEmissao)
    xCondicao = xCondicao & " AND Periodo <= " & preparaTexto(mPeriodo)
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa DESC, [Data de Emissao] DESC, Periodo DESC, [Tipo do Movimento] DESC, [Ordem da Digitacao] DESC")
    If Localizar(55) Then
        If rsMovChequeAvista.RecordCount > 1 Then
            Do Until rsMovChequeAvista.EOF
                If xRegIgual = False Then
                    If rsMovChequeAvista("Data de Emissao").Value = mDataEmissao And rsMovChequeAvista("Tipo do Movimento").Value = mTipoMovimento And rsMovChequeAvista("Periodo").Value = mPeriodo And rsMovChequeAvista("Ordem da Digitacao").Value = mOrdemDigitacao Then
                        xRegIgual = True
                    End If
                Else
                    LocalizarAnterior = True
                    Exit Do
                End If
                rsMovChequeAvista.MoveNext
            Loop
        End If
    End If
    If LocalizarAnterior Then
        AtribuiValor
    End If
    rsMovChequeAvista.Close
    Set rsMovChequeAvista = Nothing
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
    rsMovChequeAvista.Close
    Set rsMovChequeAvista = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa ASC, [Data de Emissao] ASC, Periodo ASC, [Tipo do Movimento] ASC, [Ordem da Digitacao] ASC")
    LocalizarPrimeiro = Localizar(1)
    If LocalizarPrimeiro Then
        AtribuiValor
    End If
    rsMovChequeAvista.Close
    Set rsMovChequeAvista = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarProximo() As Boolean
Dim xCondicao As String
Dim xRegIgual As Boolean

On Error GoTo trata_erro

    LocalizarProximo = False
    xRegIgual = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND [Data de Emissao] >= " & preparaData(mDataEmissao)
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa ASC, [Data de Emissao] ASC, Periodo ASC, [Tipo do Movimento] ASC, [Ordem da Digitacao] ASC")
    If Localizar(55) Then
        If rsMovChequeAvista.RecordCount > 1 Then
            Do Until rsMovChequeAvista.EOF
                If xRegIgual = False Then
                    If rsMovChequeAvista("Data de Emissao").Value = mDataEmissao And rsMovChequeAvista("Tipo do Movimento").Value = mTipoMovimento And rsMovChequeAvista("Periodo").Value = mPeriodo And rsMovChequeAvista("Ordem da Digitacao").Value = mOrdemDigitacao Then
                        xRegIgual = True
                    End If
                Else
                    LocalizarProximo = True
                    Exit Do
                End If
                rsMovChequeAvista.MoveNext
            Loop
        End If
    End If
    If LocalizarProximo Then
        AtribuiValor
    End If
    rsMovChequeAvista.Close
    Set rsMovChequeAvista = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal nEmpresa As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Empresa = " & nEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa DESC, [Data de Emissao] DESC, Periodo DESC, [Tipo do Movimento] DESC, [Ordem da Digitacao] DESC")
    LocalizarUltimo = Localizar(1)
    If LocalizarUltimo Then
        AtribuiValor
    End If
    rsMovChequeAvista.Close
    Set rsMovChequeAvista = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    mOrdemDigitacao = ProximoRegistro(mEmpresa)
    gSQL = "INSERT INTO Movimento_Cheque_Avista ( Empresa, [Data de Emissao], Periodo, [Tipo do Movimento], Valor, "
    gSQL = gSQL & "[Ordem da Digitacao], [Codigo de Barra 1], [Codigo de Barra 2], [Codigo de Barra 3] ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlData(1, mDataEmissao, ", ")
    Call sqlTexto(1, mPeriodo, ", ")
    Call sqlTexto(1, mTipoMovimento, ", ")
    Call sqlValor(1, mValor, ", ")
    Call sqlNumero(1, mOrdemDigitacao, ", ")
    Call sqlTexto(1, mCodigoBarra1, ", ")
    Call sqlTexto(1, mCodigoBarra2, ", ")
    Call sqlTexto(1, mCodigoBarra3, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal nEmpresa As Integer, ByVal nDataEmissao As Date, ByVal nPeriodo As String, ByVal nTipoMovimento As String, ByVal nOrdemDigitacao As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Movimento_Cheque_Avista SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlData(2, ", [Data de Emissao] = ", mDataEmissao)
    Call sqlTexto(2, ", Periodo = ", mPeriodo)
    Call sqlTexto(2, ", [Tipo do Movimento] = ", mTipoMovimento)
    Call sqlValor(2, ", Valor = ", mValor)
    Call sqlNumero(2, ", [Ordem da Digitacao] = ", mOrdemDigitacao)
    Call sqlTexto(2, ", [Codigo de Barra 1] = ", mCodigoBarra1)
    Call sqlTexto(2, ", [Codigo de Barra 2] = ", mCodigoBarra2)
    Call sqlTexto(2, ", [Codigo de Barra 3] = ", mCodigoBarra3)
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND [Data de Emissao] = " & preparaData(nDataEmissao)
    gSQL = gSQL & " AND Periodo = " & preparaTexto(nPeriodo)
    gSQL = gSQL & " AND [Tipo do Movimento] = " & preparaTexto(nTipoMovimento)
    gSQL = gSQL & " AND [Ordem da Digitacao] = " & nOrdemDigitacao
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal nEmpresa As Integer, ByVal nDataEmissao As Date, ByVal nPeriodo As String, ByVal nTipoMovimento As String, ByVal nOrdemDigitacao As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM Movimento_Cheque_Avista"
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND [Data de Emissao] = " & preparaData(nDataEmissao)
    gSQL = gSQL & " AND Periodo = " & preparaTexto(nPeriodo)
    gSQL = gSQL & " AND [Tipo do Movimento] = " & preparaTexto(nTipoMovimento)
    gSQL = gSQL & " AND [Ordem da Digitacao] = " & nOrdemDigitacao
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, [Data de Emissao], Periodo, [Tipo do Movimento], Valor, [Ordem da Digitacao], [Codigo de Barra 1], [Codigo de Barra 2], [Codigo de Barra 3] FROM Movimento_Cheque_Avista" & " " & xCondicao & " " & xTipoOrdem)
End Function


Public Function ProximoRegistro(ByVal nEmpresa As Integer) As Long

On Error GoTo Err_ProximoRegistro

    'Pega o número da última ordem do movimento
    ProximoRegistro = 1
    gSQL = ""
    gSQL = gSQL & "SELECT [Ordem da Digitacao]"
    gSQL = gSQL & "     FROM Movimento_Cheque_Avista"
    gSQL = gSQL & "    WHERE Empresa = " & nEmpresa
    gSQL = gSQL & "      AND [Data de Emissao] = " & preparaData(mDataEmissao)
    gSQL = gSQL & "      AND Periodo = " & preparaTexto(mPeriodo)
    gSQL = gSQL & "      AND [Tipo do Movimento] = " & preparaTexto(mTipoMovimento)
    gSQL = gSQL & " ORDER BY [Ordem da Digitacao] DESC"
    If Localizar(1) Then
        ProximoRegistro = rsMovChequeAvista("Ordem da Digitacao").Value + 1
    End If
    rsMovChequeAvista.Close
    Set rsMovChequeAvista = Nothing
    Exit Function

Err_ProximoRegistro:
End Function


Public Function TotalPeriodo(ByVal nEmpresa As Integer, ByVal nDataEmissao As Date, ByVal nPeriodo As String, ByVal nTipoMovimento As String) As Currency

On Error GoTo trata_erro
    Dim xSQL As String
    Dim xOrdem As String
    TotalPeriodo = 0
    gSQL = "SELECT SUM(Valor) AS Total FROM Movimento_Cheque_Avista"
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND [Data de Emissao] = " & preparaData(nDataEmissao)
    gSQL = gSQL & " AND Periodo = " & preparaTexto(nPeriodo)
    gSQL = gSQL & " AND [Tipo do Movimento] = " & preparaTexto(nTipoMovimento)
    If Localizar(0) Then
        If Not IsNull(rsMovChequeAvista("Total").Value) Then
            TotalPeriodo = rsMovChequeAvista("Total").Value
        End If
    End If
    rsMovChequeAvista.Close
    Set rsMovChequeAvista = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function
'Fim Métodos da Classe


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsMovChequeAvista = New adodb.Recordset
    rsMovChequeAvista.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsMovChequeAvista.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsMovChequeAvista.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsMovChequeAvista("Empresa").Value
    DataEmissao = rsMovChequeAvista("Data de Emissao").Value
    Periodo = rsMovChequeAvista("Periodo").Value
    TipoMovimento = rsMovChequeAvista("Tipo do Movimento").Value
    Valor = rsMovChequeAvista("Valor").Value
    OrdemDigitacao = rsMovChequeAvista("Ordem da Digitacao").Value
    CodigoBarra1 = rsMovChequeAvista("Codigo de Barra 1").Value
    CodigoBarra2 = rsMovChequeAvista("Codigo de Barra 2").Value
    CodigoBarra3 = rsMovChequeAvista("Codigo de Barra 3").Value
End Sub


Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function

