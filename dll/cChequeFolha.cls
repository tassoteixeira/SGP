VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cChequeFolha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mNumero As String
Private mData As Date
Private mValor As Currency
Private mNome As String
Private mHistorico As String
Private mSituacao As String
Private mDataVencimento As Date
Private mDataBaixa As Date
Private mTipoCheque As Integer
Private mNumeroConta As String
Private mChequeemPosse As Integer

Private rsChequeFolha As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get Empresa() As Integer
    Empresa = mEmpresa
End Property
Public Property Let Empresa(ByVal Valor As Integer)
    mEmpresa = Valor
End Property


Public Property Get Numero() As String
    Numero = mNumero
End Property
Public Property Let Numero(ByVal Valor As String)
    mNumero = Valor
End Property


Public Property Get Data() As Date
    Data = mData
End Property
Public Property Let Data(ByVal Valor As Date)
    mData = Valor
End Property


Public Property Get Valor() As Currency
    Valor = mValor
End Property
Public Property Let Valor(ByVal Valor As Currency)
    mValor = Valor
End Property


Public Property Get Nome() As String
    Nome = mNome
End Property
Public Property Let Nome(ByVal Valor As String)
    mNome = Valor
End Property


Public Property Get Historico() As String
    Historico = mHistorico
End Property
Public Property Let Historico(ByVal Valor As String)
    mHistorico = Valor
End Property


Public Property Get Situacao() As String
    Situacao = mSituacao
End Property
Public Property Let Situacao(ByVal Valor As String)
    mSituacao = Valor
End Property


Public Property Get DataVencimento() As Date
    DataVencimento = mDataVencimento
End Property
Public Property Let DataVencimento(ByVal Valor As Date)
    mDataVencimento = Valor
End Property


Public Property Get DataBaixa() As Date
    DataBaixa = mDataBaixa
End Property
Public Property Let DataBaixa(ByVal Valor As Date)
    mDataBaixa = Valor
End Property


Public Property Get TipoCheque() As Integer
    TipoCheque = mTipoCheque
End Property
Public Property Let TipoCheque(ByVal Valor As Integer)
    mTipoCheque = Valor
End Property


Public Property Get NumeroConta() As String
    NumeroConta = mNumeroConta
End Property
Public Property Let NumeroConta(ByVal Valor As String)
    mNumeroConta = Valor
End Property


Public Property Get ChequeemPosse() As Integer
    ChequeemPosse = mChequeemPosse
End Property
Public Property Let ChequeemPosse(ByVal Valor As Integer)
    mChequeemPosse = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pNumero As String) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND Data = " & preparaData(pData)
    xCondicao = xCondicao & " AND Numero = " & preparaTexto(pNumero)
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo = True Then
        AtribuiValor
    End If
    rsChequeFolha.Close
    Set rsChequeFolha = Nothing
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
    gSQL = PreparaSQL(xSQL, "ORDER BY Data DESC, Numero DESC")
    
    If Localizar(55) Then
        If rsChequeFolha.RecordCount > 1 Then
            Do Until rsChequeFolha.EOF
                If xRegIgual = False Then
                    If rsChequeFolha("Data").Value = mData And rsChequeFolha("Numero").Value = mNumero Then
                        xRegIgual = True
                    End If
                Else
                    LocalizarAnterior = True
                    AtribuiValor
                    Exit Do
                End If
                rsChequeFolha.MoveNext
            Loop
        End If
    End If
    rsChequeFolha.Close
    Set rsChequeFolha = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY Data ASC, Numero ASC")
    If Localizar(1) Then
        LocalizarPrimeiro = True
        AtribuiValor
    End If
    rsChequeFolha.Close
    Set rsChequeFolha = Nothing
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
    gSQL = PreparaSQL(xSQL, "ORDER BY Data ASC, Numero ASC")
    
    If Localizar(55) Then
        If rsChequeFolha.RecordCount > 1 Then
            Do Until rsChequeFolha.EOF
                If xRegIgual = False Then
                    If rsChequeFolha("Data").Value = mData And rsChequeFolha("Numero").Value = mNumero Then
                        xRegIgual = True
                    End If
                Else
                    LocalizarProximo = True
                    AtribuiValor
                    Exit Do
                End If
                rsChequeFolha.MoveNext
            Loop
        End If
    End If
    rsChequeFolha.Close
    Set rsChequeFolha = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal pEmpresa As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY Data DESC, Numero DESC")
    If Localizar(1) Then
        LocalizarUltimo = True
        AtribuiValor
    End If
    rsChequeFolha.Close
    Set rsChequeFolha = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Cheque_Folha ( Empresa, Numero, Data, Valor, Nome, "
    gSQL = gSQL & "Historico, Situacao, [Data de Vencimento], [Data da Baixa], [Tipo de Cheque], "
    gSQL = gSQL & "[Numero da Conta], [Cheque em Posse] ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlTexto(1, mNumero, ", ")
    Call sqlData(1, mData, ", ")
    Call sqlValor(1, mValor, ", ")
    Call sqlTexto(1, mNome, ", ")
    Call sqlTexto(1, mHistorico, ", ")
    Call sqlTexto(1, mSituacao, ", ")
    Call sqlData(1, mDataVencimento, ", ")
    Call sqlData(1, mDataBaixa, ", ")
    Call sqlNumero(1, mTipoCheque, ", ")
    Call sqlTexto(1, mNumeroConta, ", ")
    Call sqlNumero(1, mChequeemPosse, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pNumero As String) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Cheque_Folha SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlTexto(2, ", Numero = ", mNumero)
    Call sqlData(2, ", Data = ", mData)
    Call sqlValor(2, ", Valor = ", mValor)
    Call sqlTexto(2, ", Nome = ", mNome)
    Call sqlTexto(2, ", Historico = ", mHistorico)
    Call sqlTexto(2, ", Situacao = ", mSituacao)
    Call sqlData(2, ", [Data de Vencimento] = ", mDataVencimento)
    Call sqlData(2, ", [Data da Baixa] = ", mDataBaixa)
    Call sqlNumero(2, ", [Tipo de Cheque] = ", mTipoCheque)
    Call sqlTexto(2, ", [Numero da Conta] = ", mNumeroConta)
    Call sqlNumero(2, ", [Cheque em Posse] = ", mChequeemPosse)
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND Numero = " & preparaTexto(pNumero)
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pNumero As String) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM Cheque_Folha"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND Numero = " & preparaTexto(pNumero)
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


Public Function ProximoNumeroCheque(ByVal pEmpresa As Integer, ByVal pNumeroConta As String) As Long

On Error GoTo Err_ProximoNumeroCheque

    ProximoNumeroCheque = 1
    gSQL = ""
    gSQL = gSQL & "SELECT Numero"
    gSQL = gSQL & "  FROM Cheque_Folha"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Numero da Conta] = " & preparaTexto(pNumeroConta)
    gSQL = gSQL & " ORDER BY Numero DESC"
    If Localizar(1) Then
        If Not IsNull(rsChequeFolha("Numero").Value) Then
            ProximoNumeroCheque = CLng(Trim(rsChequeFolha("Numero").Value)) + 1
        End If
    End If
    rsChequeFolha.Close
    Set rsChequeFolha = Nothing
    Exit Function

Err_ProximoNumeroCheque:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, Numero, Data, Valor, Nome, Historico, Situacao, [Data de Vencimento], [Data da Baixa], [Tipo de Cheque], [Numero da Conta], [Cheque em Posse] FROM Cheque_Folha" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal pQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsChequeFolha = New adodb.Recordset
    rsChequeFolha.CursorLocation = adUseClient
    i = Len(gSQL)
    If pQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & pQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsChequeFolha.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsChequeFolha.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsChequeFolha("Empresa").Value
    Numero = rsChequeFolha("Numero").Value
    Data = rsChequeFolha("Data").Value
    Valor = rsChequeFolha("Valor").Value
    Nome = rsChequeFolha("Nome").Value
    Historico = rsChequeFolha("Historico").Value
    Situacao = rsChequeFolha("Situacao").Value
    If IsNull(rsChequeFolha("Data de Vencimento").Value) Then
        DataVencimento = "00:00:00"
    Else
        DataVencimento = rsChequeFolha("Data de Vencimento").Value
    End If
    If IsNull(rsChequeFolha("Data da Baixa").Value) Then
        DataBaixa = "00:00:00"
    Else
        DataBaixa = rsChequeFolha("Data da Baixa").Value
    End If
    TipoCheque = rsChequeFolha("Tipo de Cheque").Value
    NumeroConta = rsChequeFolha("Numero da Conta").Value
    ChequeemPosse = rsChequeFolha("Cheque em Posse").Value
End Sub


Private Function Localizar(ByVal pQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(pQtdRegistro) Then
        Localizar = True
    End If
End Function

