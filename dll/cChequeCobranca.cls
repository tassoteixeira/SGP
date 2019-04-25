VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cChequeCobranca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mNumeroBanco As Integer
Private mNumeroAgencia As String
Private mNumeroConta As String
Private mNumeroCheque As String
Private mValor As Currency
Private mEmitente As String
Private mMotivo As String

Private rsChequeCobranca As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get Empresa() As Integer
    Empresa = mEmpresa
End Property
Public Property Let Empresa(ByVal Valor As Integer)
    mEmpresa = Valor
End Property


Public Property Get NumeroBanco() As Integer
    NumeroBanco = mNumeroBanco
End Property
Public Property Let NumeroBanco(ByVal Valor As Integer)
    mNumeroBanco = Valor
End Property


Public Property Get NumeroAgencia() As String
    NumeroAgencia = mNumeroAgencia
End Property
Public Property Let NumeroAgencia(ByVal Valor As String)
    mNumeroAgencia = Valor
End Property


Public Property Get NumeroConta() As String
    NumeroConta = mNumeroConta
End Property
Public Property Let NumeroConta(ByVal Valor As String)
    mNumeroConta = Valor
End Property


Public Property Get NumeroCheque() As String
    NumeroCheque = mNumeroCheque
End Property
Public Property Let NumeroCheque(ByVal Valor As String)
    mNumeroCheque = Valor
End Property


Public Property Get Valor() As Currency
    Valor = mValor
End Property
Public Property Let Valor(ByVal Valor As Currency)
    mValor = Valor
End Property


Public Property Get Emitente() As String
    Emitente = mEmitente
End Property
Public Property Let Emitente(ByVal Valor As String)
    mEmitente = Valor
End Property


Public Property Get Motivo() As String
    Motivo = mMotivo
End Property
Public Property Let Motivo(ByVal Valor As String)
    mMotivo = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal pEmpresa As Integer, ByVal pNumeroCheque As String, ByVal pNumeroBanco As Integer, ByVal pNumeroAgencia As String, ByVal pNumeroConta As String) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND [Numero do Cheque] = " & preparaTexto(pNumeroCheque)
    xCondicao = xCondicao & " AND [Numero do Banco] = " & pNumeroBanco
    xCondicao = xCondicao & " AND [Numero da Agencia] = " & preparaTexto(pNumeroAgencia)
    xCondicao = xCondicao & " AND [Numero da Conta] = " & preparaTexto(pNumeroConta)
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo = True Then
        AtribuiValor
    End If
    rsChequeCobranca.Close
    Set rsChequeCobranca = Nothing
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
    xCondicao = xCondicao & " AND [Numero do Cheque] <= " & preparaTexto(mNumeroCheque)
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Numero do Cheque] DESC, [Numero do Banco] DESC, [Numero da Agencia] DESC, [Numero da Conta] DESC")
    If Localizar(55) Then
        If rsChequeCobranca.RecordCount > 1 Then
            Do Until rsChequeCobranca.EOF
                If xRegIgual = False Then
                    If rsChequeCobranca("Numero do Cheque").Value = mNumeroCheque And rsChequeCobranca("Numero do Banco").Value = mNumeroBanco And rsChequeCobranca("Numero da Agencia").Value = mNumeroAgencia And rsChequeCobranca("Numero da Conta").Value = mNumeroConta Then
                        xRegIgual = True
                    End If
                Else
                    LocalizarAnterior = True
                    Exit Do
                End If
                rsChequeCobranca.MoveNext
            Loop
        End If
    End If
    If LocalizarAnterior = True Then
        AtribuiValor
    End If
    rsChequeCobranca.Close
    Set rsChequeCobranca = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarNome(ByVal pNome As String) As Boolean

On Error GoTo trata_erro
    
    LocalizarNome = False
    gSQL = PreparaSQL(" WHERE Nome = " & preparaTexto(pNome), "")
    LocalizarNome = Localizar(1)
    If LocalizarNome = True Then
        AtribuiValor
    End If
    rsChequeCobranca.Close
    Set rsChequeCobranca = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Numero do Cheque] ASC, [Numero do Banco] ASC, [Numero da Agencia] ASC, [Numero da Conta] ASC")
    LocalizarPrimeiro = Localizar(1)
    If LocalizarPrimeiro = True Then
        AtribuiValor
    End If
    rsChequeCobranca.Close
    Set rsChequeCobranca = Nothing
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
    xCondicao = xCondicao & " AND [Numero do Cheque] >= " & preparaTexto(mNumeroCheque)
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Numero do Cheque] ASC, [Numero do Banco] ASC, [Numero da Agencia] ASC, [Numero da Conta] ASC")
    If Localizar(55) Then
        If rsChequeCobranca.RecordCount > 1 Then
            Do Until rsChequeCobranca.EOF
                If xRegIgual = False Then
                    If rsChequeCobranca("Numero do Cheque").Value = mNumeroCheque And rsChequeCobranca("Numero do Banco").Value = mNumeroBanco And rsChequeCobranca("Numero da Agencia").Value = mNumeroAgencia And rsChequeCobranca("Numero da Conta").Value = mNumeroConta Then
                        xRegIgual = True
                    End If
                Else
                    LocalizarProximo = True
                    Exit Do
                End If
                rsChequeCobranca.MoveNext
            Loop
        End If
    End If
    If LocalizarProximo = True Then
        AtribuiValor
    End If
    rsChequeCobranca.Close
    Set rsChequeCobranca = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal pEmpresa As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Numero do Cheque] DESC, [Numero do Banco] DESC, [Numero da Agencia] DESC, [Numero da Conta] DESC")
    LocalizarUltimo = Localizar(1)
    If LocalizarUltimo = True Then
        AtribuiValor
    End If
    rsChequeCobranca.Close
    Set rsChequeCobranca = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Cheque_Cobranca ( Empresa, [Numero do Banco], [Numero da Agencia], [Numero da Conta], [Numero do Cheque], "
    gSQL = gSQL & "Valor, Emitente, Motivo ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlNumero(1, mNumeroBanco, ", ")
    Call sqlTexto(1, mNumeroAgencia, ", ")
    Call sqlTexto(1, mNumeroConta, ", ")
    Call sqlTexto(1, mNumeroCheque, ", ")
    Call sqlValor(1, mValor, ", ")
    Call sqlTexto(1, mEmitente, ", ")
    Call sqlTexto(1, mMotivo, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal pEmpresa As Integer, ByVal pNumeroCheque As String, ByVal pNumeroBanco As Integer, ByVal pNumeroAgencia As String, ByVal pNumeroConta As String) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Cheque_Cobranca SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlNumero(2, ", [Numero do Banco] = ", mNumeroBanco)
    Call sqlTexto(2, ", [Numero da Agencia] = ", mNumeroAgencia)
    Call sqlTexto(2, ", [Numero da Conta] = ", mNumeroConta)
    Call sqlTexto(2, ", [Numero do Cheque] = ", mNumeroCheque)
    Call sqlValor(2, ", Valor = ", mValor)
    Call sqlTexto(2, ", Emitente = ", mEmitente)
    Call sqlTexto(2, ", Motivo = ", mMotivo)
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Numero do Cheque] = " & preparaTexto(pNumeroCheque)
    gSQL = gSQL & " AND [Numero do Banco] = " & pNumeroBanco
    gSQL = gSQL & " AND [Numero da Agencia] = " & preparaTexto(pNumeroAgencia)
    gSQL = gSQL & " AND [Numero da Conta] = " & preparaTexto(pNumeroConta)
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal pEmpresa As Integer, ByVal pNumeroCheque As String, ByVal pNumeroBanco As Integer, ByVal pNumeroAgencia As String, ByVal pNumeroConta As String) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM Cheque_Cobranca"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Numero do Cheque] = " & preparaTexto(pNumeroCheque)
    gSQL = gSQL & " AND [Numero do Banco] = " & pNumeroBanco
    gSQL = gSQL & " AND [Numero da Agencia] = " & preparaTexto(pNumeroAgencia)
    gSQL = gSQL & " AND [Numero da Conta] = " & preparaTexto(pNumeroConta)
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, [Numero do Banco], [Numero da Agencia], [Numero da Conta], [Numero do Cheque], Valor, Emitente, Motivo FROM Cheque_Cobranca" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal pQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsChequeCobranca = New adodb.Recordset
    rsChequeCobranca.CursorLocation = adUseClient
    i = Len(gSQL)
    If pQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & pQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsChequeCobranca.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsChequeCobranca.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsChequeCobranca("Empresa").Value
    NumeroBanco = rsChequeCobranca("Numero do Banco").Value
    NumeroAgencia = rsChequeCobranca("Numero da Agencia").Value
    NumeroConta = rsChequeCobranca("Numero da Conta").Value
    NumeroCheque = rsChequeCobranca("Numero do Cheque").Value
    Valor = rsChequeCobranca("Valor").Value
    Emitente = rsChequeCobranca("Emitente").Value
    Motivo = rsChequeCobranca("Motivo").Value
End Sub


Private Function Localizar(ByVal pQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(pQtdRegistro) Then
        Localizar = True
    End If
End Function

