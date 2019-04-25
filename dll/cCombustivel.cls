VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCombustivel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mCodigo As String
Private mNome As String
Private mQuantidadeEmEstoque As Currency
Private mCapacidadeArmazenamento As Currency
Private mPrecoMedio As Currency
Private mOrdem As Integer

Private rsCombustivel As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get Empresa() As Integer
    Empresa = mEmpresa
End Property
Public Property Let Empresa(ByVal Valor As Integer)
    mEmpresa = Valor
End Property


Public Property Get Codigo() As String
    Codigo = mCodigo
End Property
Public Property Let Codigo(ByVal Valor As String)
    mCodigo = Valor
End Property


Public Property Get Nome() As String
    Nome = mNome
End Property
Public Property Let Nome(ByVal Valor As String)
    mNome = Valor
End Property


Public Property Get QuantidadeEmEstoque() As Currency
    QuantidadeEmEstoque = mQuantidadeEmEstoque
End Property
Public Property Let QuantidadeEmEstoque(ByVal Valor As Currency)
    mQuantidadeEmEstoque = Valor
End Property


Public Property Get CapacidadeArmazenamento() As Currency
    CapacidadeArmazenamento = mCapacidadeArmazenamento
End Property
Public Property Let CapacidadeArmazenamento(ByVal Valor As Currency)
    mCapacidadeArmazenamento = Valor
End Property


Public Property Get PrecoMedio() As Currency
    PrecoMedio = mPrecoMedio
End Property
Public Property Let PrecoMedio(ByVal Valor As Currency)
    mPrecoMedio = Valor
End Property


Public Property Get Ordem() As Integer
    Ordem = mOrdem
End Property
Public Property Let Ordem(ByVal Valor As Integer)
    mOrdem = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal nEmpresa As Integer, ByVal nCodigo As String) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & nEmpresa
    xCondicao = xCondicao & " AND Codigo = " & preparaTexto(nCodigo)
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarAnterior() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarAnterior = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND Codigo < " & preparaTexto(mCodigo)
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa DESC, Codigo DESC")
    LocalizarAnterior = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarNome(ByVal nEmpresa As Integer, ByVal nNome As String) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro
    
    LocalizarNome = False
    
    xCondicao = " WHERE Empresa = " & nEmpresa
    xCondicao = xCondicao & " AND Nome = " & preparaTexto(nNome)
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarNome = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro(ByVal nEmpresa As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE Empresa = " & nEmpresa
    xCondicao = xCondicao & " AND Codigo > " & preparaTexto("  ")
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa ASC, Codigo ASC")
    LocalizarPrimeiro = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarProximo() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarProximo = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND Codigo > " & preparaTexto(mCodigo)
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa ASC, Codigo ASC")
    LocalizarProximo = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal nEmpresa As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Empresa = " & nEmpresa
    xCondicao = xCondicao & " AND Codigo < " & preparaTexto("ZZ")
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa DESC, Codigo DESC")
    LocalizarUltimo = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Combustivel ( Empresa, Codigo, Nome, [Quantidade em Estoque], [Capacidade de Armazenamento], "
    gSQL = gSQL & "[Preco Medio], Ordem ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlTexto(1, mCodigo, ", ")
    Call sqlTexto(1, mNome, ", ")
    Call sqlValor(1, mQuantidadeEmEstoque, ", ")
    Call sqlValor(1, mCapacidadeArmazenamento, ", ")
    Call sqlValor(1, mPrecoMedio, ", ")
    Call sqlNumero(1, mOrdem, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal nEmpresa As Integer, ByVal nCodigo As String) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Combustivel SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlTexto(2, ", Codigo = ", mCodigo)
    Call sqlTexto(2, ", Nome = ", mNome)
    Call sqlValor(2, ", [Quantidade em Estoque] = ", mQuantidadeEmEstoque)
    Call sqlValor(2, ", [Capacidade de Armazenamento] = ", mCapacidadeArmazenamento)
    Call sqlValor(2, ", [Preco Medio] = ", mPrecoMedio)
    Call sqlNumero(2, ", Ordem = ", mOrdem)
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND Codigo = " & preparaTexto(nCodigo)
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal nEmpresa As Integer, ByVal nCodigo As String) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM Combustivel"
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND Codigo = " & preparaTexto(nCodigo)
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, Codigo, Nome, [Quantidade em Estoque], [Capacidade de Armazenamento], [Preco Medio], Ordem FROM Combustivel" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsCombustivel = New adodb.Recordset
    rsCombustivel.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsCombustivel.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsCombustivel.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsCombustivel("Empresa").Value
    Codigo = rsCombustivel("Codigo").Value
    Nome = rsCombustivel("Nome").Value
    QuantidadeEmEstoque = rsCombustivel("Quantidade em Estoque").Value
    CapacidadeArmazenamento = rsCombustivel("Capacidade de Armazenamento").Value
    PrecoMedio = rsCombustivel("Preco Medio").Value
    Ordem = rsCombustivel("Ordem").Value
End Sub


Private Function Localizar() As Boolean
    Localizar = False
    If AtualizaRecordset(1) Then
        Localizar = True
        AtribuiValor
    End If
    rsCombustivel.Close
    Set rsCombustivel = Nothing
End Function

