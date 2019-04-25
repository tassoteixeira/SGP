VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSaidaTransferenciaProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mDataTransferencia As Date
Private mCodigoProduto2 As Long
Private mNumeroDocumento As Long
Private mEntrounaEmpresa As Integer
Private mPrecoCusto As Currency
Private mQuantidade As Currency
Private mDataDigitacao As Date
Private mObservacao As String

Private rsSaidaTransferenciaProduto As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get Empresa() As Integer
    Empresa = mEmpresa
End Property
Public Property Let Empresa(ByVal Valor As Integer)
    mEmpresa = Valor
End Property


Public Property Get DataTransferencia() As Date
    DataTransferencia = mDataTransferencia
End Property
Public Property Let DataTransferencia(ByVal Valor As Date)
    mDataTransferencia = Valor
End Property


Public Property Get CodigoProduto2() As Long
    CodigoProduto2 = mCodigoProduto2
End Property
Public Property Let CodigoProduto2(ByVal Valor As Long)
    mCodigoProduto2 = Valor
End Property


Public Property Get NumeroDocumento() As Long
    NumeroDocumento = mNumeroDocumento
End Property
Public Property Let NumeroDocumento(ByVal Valor As Long)
    mNumeroDocumento = Valor
End Property


Public Property Get EntrounaEmpresa() As Integer
    EntrounaEmpresa = mEntrounaEmpresa
End Property
Public Property Let EntrounaEmpresa(ByVal Valor As Integer)
    mEntrounaEmpresa = Valor
End Property


Public Property Get PrecoCusto() As Currency
    PrecoCusto = mPrecoCusto
End Property
Public Property Let PrecoCusto(ByVal Valor As Currency)
    mPrecoCusto = Valor
End Property


Public Property Get Quantidade() As Currency
    Quantidade = mQuantidade
End Property
Public Property Let Quantidade(ByVal Valor As Currency)
    mQuantidade = Valor
End Property


Public Property Get DataDigitacao() As Date
    DataDigitacao = mDataDigitacao
End Property
Public Property Let DataDigitacao(ByVal Valor As Date)
    mDataDigitacao = Valor
End Property


Public Property Get Observacao() As String
    Observacao = mObservacao
End Property
Public Property Let Observacao(ByVal Valor As String)
    mObservacao = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal pEmpresa As Integer, ByVal pDataTransferencia As Date, ByVal pCodigoProduto2 As Long, ByVal pNumeroDocumento As Long) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND [Data da Transferencia] = " & preparaData(pDataTransferencia)
    xCondicao = xCondicao & " AND [Codigo do Produto2] = " & pCodigoProduto2
    xCondicao = xCondicao & " AND [Numero do Documento] = " & pNumeroDocumento
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo = True Then
        AtribuiValor
    End If
    rsSaidaTransferenciaProduto.Close
    Set rsSaidaTransferenciaProduto = Nothing
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
    xCondicao = xCondicao & " AND [Data da Transferencia] <= " & preparaData(mDataTransferencia)
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Data da Transferencia] DESC, [Codigo do Produto2] DESC, [Numero do Documento] DESC")
    If Localizar(55) Then
        If rsSaidaTransferenciaProduto.RecordCount > 1 Then
            Do Until rsSaidaTransferenciaProduto.EOF
                If xRegIgual = False Then
                    If rsSaidaTransferenciaProduto("Data da Transferencia").Value = mDataTransferencia And rsSaidaTransferenciaProduto("Codigo do Produto2").Value = mCodigoProduto2 And rsSaidaTransferenciaProduto("Numero do Documento").Value = mNumeroDocumento Then
                        xRegIgual = True
                    End If
                Else
                    LocalizarAnterior = True
                    Exit Do
                End If
                rsSaidaTransferenciaProduto.MoveNext
            Loop
        End If
    End If
    If LocalizarAnterior = True Then
        AtribuiValor
    End If
    rsSaidaTransferenciaProduto.Close
    Set rsSaidaTransferenciaProduto = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Data da Transferencia] ASC, [Codigo do Produto2] ASC, [Numero do Documento] ASC")
    LocalizarPrimeiro = Localizar(1)
    If LocalizarPrimeiro = True Then
        AtribuiValor
    End If
    rsSaidaTransferenciaProduto.Close
    Set rsSaidaTransferenciaProduto = Nothing
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
    xCondicao = xCondicao & " AND [Data da Transferencia] >= " & preparaData(mDataTransferencia)
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Data da Transferencia] ASC, [Codigo do Produto2] ASC, [Numero do Documento] ASC")
    If Localizar(55) Then
        If rsSaidaTransferenciaProduto.RecordCount > 1 Then
            Do Until rsSaidaTransferenciaProduto.EOF
                If xRegIgual = False Then
                    If rsSaidaTransferenciaProduto("Data da Transferencia").Value = mDataTransferencia And rsSaidaTransferenciaProduto("Codigo do Produto2").Value = mCodigoProduto2 And rsSaidaTransferenciaProduto("Numero do Documento").Value = mNumeroDocumento Then
                        xRegIgual = True
                    End If
                Else
                    LocalizarProximo = True
                    Exit Do
                End If
                rsSaidaTransferenciaProduto.MoveNext
            Loop
        End If
    End If
    If LocalizarProximo = True Then
        AtribuiValor
    End If
    rsSaidaTransferenciaProduto.Close
    Set rsSaidaTransferenciaProduto = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal pEmpresa As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Data da Transferencia] DESC, [Codigo do Produto2] DESC, [Numero do Documento] DESC")
    LocalizarUltimo = Localizar(1)
    If LocalizarUltimo = True Then
        AtribuiValor
    End If
    rsSaidaTransferenciaProduto.Close
    Set rsSaidaTransferenciaProduto = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Saida_Transferencia_Produto ( Empresa, [Data da Transferencia], [Codigo do Produto2], [Numero do Documento], [Entrou na Empresa], "
    gSQL = gSQL & "[Preco de Custo], Quantidade, [Data da Digitacao], Observacao ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlData(1, mDataTransferencia, ", ")
    Call sqlNumero(1, mCodigoProduto2, ", ")
    Call sqlNumero(1, mNumeroDocumento, ", ")
    Call sqlNumero(1, mEntrounaEmpresa, ", ")
    Call sqlValor(1, mPrecoCusto, ", ")
    Call sqlValor(1, mQuantidade, ", ")
    Call sqlData(1, mDataDigitacao, ", ")
    Call sqlTexto(1, mObservacao, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal pEmpresa As Integer, ByVal pDataTransferencia As Date, ByVal pCodigoProduto2 As Long, ByVal pNumeroDocumento As Long) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Saida_Transferencia_Produto SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlData(2, ", [Data da Transferencia] = ", mDataTransferencia)
    Call sqlNumero(2, ", [Codigo do Produto2] = ", mCodigoProduto2)
    Call sqlNumero(2, ", [Numero do Documento] = ", mNumeroDocumento)
    Call sqlNumero(2, ", [Entrou na Empresa] = ", mEntrounaEmpresa)
    Call sqlValor(2, ", [Preco de Custo] = ", mPrecoCusto)
    Call sqlValor(2, ", Quantidade = ", mQuantidade)
    Call sqlData(2, ", [Data da Digitacao] = ", mDataDigitacao)
    Call sqlTexto(2, ", Observacao = ", mObservacao)
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Data da Transferencia] = " & preparaData(pDataTransferencia)
    gSQL = gSQL & " AND [Codigo do Produto2] = " & pCodigoProduto2
    gSQL = gSQL & " AND [Numero do Documento] = " & pNumeroDocumento
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal pEmpresa As Integer, ByVal pDataTransferencia As Date, ByVal pCodigoProduto2 As Long, ByVal pNumeroDocumento As Long) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM Saida_Transferencia_Produto"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Data da Transferencia] = " & preparaData(pDataTransferencia)
    gSQL = gSQL & " AND [Codigo do Produto2] = " & pCodigoProduto2
    gSQL = gSQL & " AND [Numero do Documento] = " & pNumeroDocumento
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, [Data da Transferencia], [Codigo do Produto2], [Numero do Documento], [Entrou na Empresa], [Preco de Custo], Quantidade, [Data da Digitacao], Observacao FROM Saida_Transferencia_Produto" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal pQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsSaidaTransferenciaProduto = New adodb.Recordset
    rsSaidaTransferenciaProduto.CursorLocation = adUseClient
    i = Len(gSQL)
    If pQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & pQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsSaidaTransferenciaProduto.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsSaidaTransferenciaProduto.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsSaidaTransferenciaProduto("Empresa").Value
    DataTransferencia = rsSaidaTransferenciaProduto("Data da Transferencia").Value
    CodigoProduto2 = rsSaidaTransferenciaProduto("Codigo do Produto2").Value
    NumeroDocumento = rsSaidaTransferenciaProduto("Numero do Documento").Value
    EntrounaEmpresa = rsSaidaTransferenciaProduto("Entrou na Empresa").Value
    PrecoCusto = rsSaidaTransferenciaProduto("Preco de Custo").Value
    Quantidade = rsSaidaTransferenciaProduto("Quantidade").Value
    DataDigitacao = rsSaidaTransferenciaProduto("Data da Digitacao").Value
    Observacao = rsSaidaTransferenciaProduto("Observacao").Value
End Sub


Private Function Localizar(ByVal pQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(pQtdRegistro) Then
        Localizar = True
    End If
End Function

