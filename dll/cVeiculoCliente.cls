VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cVeiculoCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCodigoCliente As Long
Private mCodigoVeiculo As Integer
Private mNome As String
Private mAno As Integer
Private mCor As String
Private mPlacaLetra As String
Private mPlacaNumero As Integer

Private rsVeiculoCliente As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get CodigoCliente() As Long
    CodigoCliente = mCodigoCliente
End Property
Public Property Let CodigoCliente(ByVal Valor As Long)
    mCodigoCliente = Valor
End Property


Public Property Get CodigoVeiculo() As Integer
    CodigoVeiculo = mCodigoVeiculo
End Property
Public Property Let CodigoVeiculo(ByVal Valor As Integer)
    mCodigoVeiculo = Valor
End Property


Public Property Get Nome() As String
    Nome = mNome
End Property
Public Property Let Nome(ByVal Valor As String)
    mNome = Valor
End Property


Public Property Get Ano() As Integer
    Ano = mAno
End Property
Public Property Let Ano(ByVal Valor As Integer)
    mAno = Valor
End Property


Public Property Get Cor() As String
    Cor = mCor
End Property
Public Property Let Cor(ByVal Valor As String)
    mCor = Valor
End Property


Public Property Get PlacaLetra() As String
    PlacaLetra = mPlacaLetra
End Property
Public Property Let PlacaLetra(ByVal Valor As String)
    mPlacaLetra = Valor
End Property


Public Property Get PlacaNumero() As Integer
    PlacaNumero = mPlacaNumero
End Property
Public Property Let PlacaNumero(ByVal Valor As Integer)
    mPlacaNumero = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function ClienteTemVeiculo(ByVal pCodigoCliente As Long) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    ClienteTemVeiculo = False
    xCondicao = " WHERE [Codigo do Cliente] = " & pCodigoCliente
    gSQL = PreparaSQL(xCondicao, "")
    If Localizar(1) Then
        ClienteTemVeiculo = True
    End If
    rsVeiculoCliente.Close
    Set rsVeiculoCliente = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarCodigo(ByVal pCodigoCliente As Long, ByVal pCodigoVeiculo As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE [Codigo do Cliente] = " & pCodigoCliente
    xCondicao = xCondicao & " AND [Codigo do Veiculo] = " & pCodigoVeiculo
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo = True Then
        AtribuiValor
    End If
    rsVeiculoCliente.Close
    Set rsVeiculoCliente = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarAnterior() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarAnterior = False
    xCondicao = " WHERE [Codigo do Cliente] <= " & mCodigoCliente
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Codigo do Cliente] DESC, [Codigo do Veiculo] DESC")
    LocalizarAnterior = Localizar(1)
    If LocalizarAnterior = True Then
        AtribuiValor
    End If
    rsVeiculoCliente.Close
    Set rsVeiculoCliente = Nothing
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
    rsVeiculoCliente.Close
    Set rsVeiculoCliente = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    gSQL = PreparaSQL("", "ORDER BY [Codigo do Cliente] ASC, [Codigo do Veiculo] ASC")
    LocalizarPrimeiro = Localizar(1)
    If LocalizarPrimeiro = True Then
        AtribuiValor
    End If
    rsVeiculoCliente.Close
    Set rsVeiculoCliente = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarProximo() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarProximo = False
    xCondicao = " WHERE [Codigo do Cliente] >= " & mCodigoCliente
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Codigo do Cliente] ASC, [Codigo do Veiculo] ASC")
    LocalizarProximo = Localizar(1)
    If LocalizarProximo = True Then
        AtribuiValor
    End If
    rsVeiculoCliente.Close
    Set rsVeiculoCliente = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo() As Boolean
On Error GoTo trata_erro

    LocalizarUltimo = False
    gSQL = PreparaSQL("", "ORDER BY [Codigo do Cliente] DESC, [Codigo do Veiculo] DESC")
    LocalizarUltimo = Localizar(1)
    If LocalizarUltimo = True Then
        AtribuiValor
    End If
    rsVeiculoCliente.Close
    Set rsVeiculoCliente = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO VeiculoCliente ( [Codigo do Cliente], [Codigo do Veiculo], Nome, Ano, Cor, "
    gSQL = gSQL & "[Placa Letra], [Placa Numero] ) VALUES ( "
    Call sqlNumero(1, mCodigoCliente, ", ")
    Call sqlNumero(1, mCodigoVeiculo, ", ")
    Call sqlTexto(1, mNome, ", ")
    Call sqlNumero(1, mAno, ", ")
    Call sqlTexto(1, mCor, ", ")
    Call sqlTexto(1, mPlacaLetra, ", ")
    Call sqlNumero(1, mPlacaNumero, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal pCodigoCliente As Long, ByVal pCodigoVeiculo As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE VeiculoCliente SET "
    Call sqlNumero(2, "[Codigo do Cliente] = ", mCodigoCliente)
    Call sqlNumero(2, ", [Codigo do Veiculo] = ", mCodigoVeiculo)
    Call sqlTexto(2, ", Nome = ", mNome)
    Call sqlNumero(2, ", Ano = ", mAno)
    Call sqlTexto(2, ", Cor = ", mCor)
    Call sqlTexto(2, ", [Placa Letra] = ", mPlacaLetra)
    Call sqlNumero(2, ", [Placa Numero] = ", mPlacaNumero)
    gSQL = gSQL & " WHERE [Codigo do Cliente] = " & pCodigoCliente
    gSQL = gSQL & "   AND [Codigo do Veiculo] = " & pCodigoVeiculo
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal pCodigoCliente As Long, ByVal pCodigoVeiculo As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM VeiculoCliente"
    gSQL = gSQL & " WHERE [Codigo do Cliente] = " & pCodigoCliente
    gSQL = gSQL & "   AND [Codigo do Veiculo] = " & pCodigoVeiculo
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT [Codigo do Cliente], [Codigo do Veiculo], Nome, Ano, Cor, [Placa Letra], [Placa Numero] FROM VeiculoCliente" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal pQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsVeiculoCliente = New adodb.Recordset
    rsVeiculoCliente.CursorLocation = adUseClient
    i = Len(gSQL)
    If pQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & pQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsVeiculoCliente.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsVeiculoCliente.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    CodigoCliente = rsVeiculoCliente("Codigo do Cliente").Value
    CodigoVeiculo = rsVeiculoCliente("Codigo do Veiculo").Value
    Nome = rsVeiculoCliente("Nome").Value
    Ano = rsVeiculoCliente("Ano").Value
    Cor = rsVeiculoCliente("Cor").Value
    PlacaLetra = rsVeiculoCliente("Placa Letra").Value
    PlacaNumero = rsVeiculoCliente("Placa Numero").Value
End Sub


Private Function Localizar(ByVal pQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(pQtdRegistro) Then
        Localizar = True
    End If
End Function

