VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAcertoVendaECF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mData As Date
Private mTipoCombustivel As String
Private mOperacao As String
Private mPrecoVenda As Currency
Private mQuantidade As Currency
Private mValorTotal As Currency
Private mObservacao As String

Private rsAcertoVendaECF As New adodb.Recordset


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


Public Property Get TipoCombustivel() As String
    TipoCombustivel = mTipoCombustivel
End Property
Public Property Let TipoCombustivel(ByVal Valor As String)
    mTipoCombustivel = Valor
End Property


Public Property Get Operacao() As String
    Operacao = mOperacao
End Property
Public Property Let Operacao(ByVal Valor As String)
    mOperacao = Valor
End Property


Public Property Get PrecoVenda() As Currency
    PrecoVenda = mPrecoVenda
End Property
Public Property Let PrecoVenda(ByVal Valor As Currency)
    mPrecoVenda = Valor
End Property


Public Property Get Quantidade() As Currency
    Quantidade = mQuantidade
End Property
Public Property Let Quantidade(ByVal Valor As Currency)
    mQuantidade = Valor
End Property


Public Property Get ValorTotal() As Currency
    ValorTotal = mValorTotal
End Property
Public Property Let ValorTotal(ByVal Valor As Currency)
    mValorTotal = Valor
End Property


Public Property Get Observacao() As String
    Observacao = mObservacao
End Property
Public Property Let Observacao(ByVal Valor As String)
    mObservacao = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pTipoCombustivel As String) As Boolean

Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND Data = " & preparaData(pData)
    xCondicao = xCondicao & " AND [Tipo de Combustivel] = " & preparaTexto(pTipoCombustivel)
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo = True Then
        AtribuiValor
    End If
    rsAcertoVendaECF.Close
    Set rsAcertoVendaECF = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarAnterior() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarAnterior = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND Data < " & preparaData(mData)
    gSQL = PreparaSQL(xCondicao, "ORDER BY Data DESC")
    LocalizarAnterior = Localizar(1)
    If LocalizarAnterior = True Then
        AtribuiValor
    End If
    rsAcertoVendaECF.Close
    Set rsAcertoVendaECF = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY Data ASC")
    LocalizarPrimeiro = Localizar(1)
    If LocalizarPrimeiro = True Then
        AtribuiValor
    End If
    rsAcertoVendaECF.Close
    Set rsAcertoVendaECF = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarProximo() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarProximo = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND Data > " & preparaData(mData)
    gSQL = PreparaSQL(xCondicao, "ORDER BY Data ASC")
    LocalizarProximo = Localizar(1)
    If LocalizarProximo = True Then
        AtribuiValor
    End If
    rsAcertoVendaECF.Close
    Set rsAcertoVendaECF = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal pEmpresa As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY Data DESC")
    LocalizarUltimo = Localizar(1)
    If LocalizarUltimo = True Then
        AtribuiValor
    End If
    rsAcertoVendaECF.Close
    Set rsAcertoVendaECF = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO AcertoVendaECF ( Empresa, Data, [Tipo de Combustivel], Operacao, Quantidade, "
    gSQL = gSQL & "Observacao ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlData(1, mData, ", ")
    Call sqlTexto(1, mTipoCombustivel, ", ")
    Call sqlTexto(1, mOperacao, ", ")
    Call sqlValor(1, mPrecoVenda, ", ")
    Call sqlValor(1, mQuantidade, ", ")
    Call sqlValor(1, mValorTotal, ", ")
    Call sqlTexto(1, mObservacao, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pTipoCombustivel As String) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE AcertoVendaECF SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlData(2, ", Data = ", mData)
    Call sqlTexto(2, ", [Tipo de Combustivel] = ", mTipoCombustivel)
    Call sqlTexto(2, ", Operacao = ", mOperacao)
    Call sqlValor(2, ", [Preco de Venda] = ", mPrecoVenda)
    Call sqlValor(2, ", Quantidade = ", mQuantidade)
    Call sqlValor(2, ", [Valor Total] = ", mValorTotal)
    Call sqlTexto(2, ", Observacao = ", mObservacao)
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND [Tipo de Combustivel] = " & preparaTexto(pTipoCombustivel)
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pTipoCombustivel As String) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM AcertoVendaECF"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND [Tipo de Combustivel] = " & preparaTexto(pTipoCombustivel)
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, Data, [Tipo de Combustivel], Operacao, [Preco de Venda], Quantidade, [Valor Total], Observacao FROM AcertoVendaECF" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal pQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsAcertoVendaECF = New adodb.Recordset
    rsAcertoVendaECF.CursorLocation = adUseClient
    i = Len(gSQL)
    If pQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & pQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsAcertoVendaECF.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsAcertoVendaECF.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsAcertoVendaECF("Empresa").Value
    Data = rsAcertoVendaECF("Data").Value
    TipoCombustivel = rsAcertoVendaECF("Tipo de Combustivel").Value
    Operacao = rsAcertoVendaECF("Operacao").Value
    PrecoVenda = rsAcertoVendaECF("Preco de Venda").Value
    Quantidade = rsAcertoVendaECF("Quantidade").Value
    ValorTotal = rsAcertoVendaECF("Valor Total").Value
    Observacao = rsAcertoVendaECF("Observacao").Value
End Sub


Private Function Localizar(ByVal pQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(pQtdRegistro) Then
        Localizar = True
    End If
End Function

