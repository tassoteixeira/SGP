VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTabelaProventoDesconto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCodigo As Integer
Private mNome As String
Private mPercentual As Currency
Private mValor As Currency
Private mFracao As String
Private mProventoouDesconto As String
Private mBaseparaCalculo As String
Private mAutomatico As Boolean

Private rsTabelaProventoDesconto As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get Codigo() As Integer
    Codigo = mCodigo
End Property
Public Property Let Codigo(ByVal Valor As Integer)
    mCodigo = Valor
End Property


Public Property Get Nome() As String
    Nome = mNome
End Property
Public Property Let Nome(ByVal Valor As String)
    mNome = Valor
End Property


Public Property Get Percentual() As Currency
    Percentual = mPercentual
End Property
Public Property Let Percentual(ByVal Valor As Currency)
    mPercentual = Valor
End Property


Public Property Get Valor() As Currency
    Valor = mValor
End Property
Public Property Let Valor(ByVal Valor As Currency)
    mValor = Valor
End Property


Public Property Get Fracao() As String
    Fracao = mFracao
End Property
Public Property Let Fracao(ByVal Valor As String)
    mFracao = Valor
End Property


Public Property Get ProventoouDesconto() As String
    ProventoouDesconto = mProventoouDesconto
End Property
Public Property Let ProventoouDesconto(ByVal Valor As String)
    mProventoouDesconto = Valor
End Property


Public Property Get BaseparaCalculo() As String
    BaseparaCalculo = mBaseparaCalculo
End Property
Public Property Let BaseparaCalculo(ByVal Valor As String)
    mBaseparaCalculo = Valor
End Property


Public Property Get Automatico() As Boolean
    Automatico = mAutomatico
End Property
Public Property Let Automatico(ByVal Valor As Boolean)
    mAutomatico = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal nCodigo As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Codigo = " & nCodigo
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
    xCondicao = " WHERE Codigo < " & mCodigo
    gSQL = PreparaSQL(xCondicao, "ORDER BY Codigo DESC")
    LocalizarAnterior = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarNome(ByVal nNome As String) As Boolean

On Error GoTo trata_erro
    
    LocalizarNome = False
    gSQL = PreparaSQL(" WHERE Nome = " & preparaTexto(nNome), "")
    LocalizarNome = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE Codigo > 0"
    gSQL = PreparaSQL(xCondicao, "ORDER BY Codigo ASC")
    LocalizarPrimeiro = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarProximo() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarProximo = False
    xCondicao = " WHERE Codigo > " & mCodigo
    gSQL = PreparaSQL(xCondicao, "ORDER BY Codigo ASC")
    LocalizarProximo = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Codigo < " & 9999
    gSQL = PreparaSQL(xCondicao, "ORDER BY Codigo DESC")
    LocalizarUltimo = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Tabela_Provento_Desconto ( Codigo, Nome, Percentual, Valor, Fracao, "
    gSQL = gSQL & "[Provento ou Desconto], [Base para Calculo], Automatico ) VALUES ( "
    Call sqlNumero(1, mCodigo, ", ")
    Call sqlTexto(1, mNome, ", ")
    Call sqlValor(1, mPercentual, ", ")
    Call sqlValor(1, mValor, ", ")
    Call sqlTexto(1, mFracao, ", ")
    Call sqlTexto(1, mProventoouDesconto, ", ")
    Call sqlTexto(1, mBaseparaCalculo, ", ")
    Call sqlBoolean(1, mAutomatico, " )")
    gConn.Execute gSQL, Incluir, adCmdText + adExecuteNoRecords
    Incluir = True
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal nCodigo As Integer) As Boolean

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Tabela_Provento_Desconto SET "
    Call sqlNumero(2, "Codigo = ", mCodigo)
    Call sqlTexto(2, ", Nome = ", mNome)
    Call sqlValor(2, ", Percentual = ", mPercentual)
    Call sqlValor(2, ", Valor = ", mValor)
    Call sqlTexto(2, ", Fracao = ", mFracao)
    Call sqlTexto(2, ", [Provento ou Desconto] = ", mProventoouDesconto)
    Call sqlTexto(2, ", [Base para Calculo] = ", mBaseparaCalculo)
    Call sqlBoolean(2, ", Automatico = ", mAutomatico)
    gSQL = gSQL & " WHERE Codigo = " & nCodigo
    gConn.Execute gSQL, Alterar, adCmdText + adExecuteNoRecords
    Alterar = True
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal nCodigo As Integer) As Boolean

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM Tabela_Provento_Desconto"
    gSQL = gSQL & " WHERE Codigo = " & nCodigo
    gConn.Execute gSQL, Excluir, adCmdText + adExecuteNoRecords
    Excluir = True
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Codigo, Nome, Percentual, Valor, Fracao, [Provento ou Desconto], [Base para Calculo], Automatico FROM Tabela_Provento_Desconto" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsTabelaProventoDesconto = New adodb.Recordset
    rsTabelaProventoDesconto.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsTabelaProventoDesconto.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsTabelaProventoDesconto.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Codigo = rsTabelaProventoDesconto("Codigo").Value
    Nome = rsTabelaProventoDesconto("Nome").Value
    Percentual = rsTabelaProventoDesconto("Percentual").Value
    Valor = rsTabelaProventoDesconto("Valor").Value
    Fracao = rsTabelaProventoDesconto("Fracao").Value
    ProventoouDesconto = rsTabelaProventoDesconto("Provento ou Desconto").Value
    BaseparaCalculo = rsTabelaProventoDesconto("Base para Calculo").Value
    Automatico = rsTabelaProventoDesconto("Automatico").Value
End Sub


Private Function Localizar() As Boolean
    Localizar = False
    If AtualizaRecordset(1) Then
        Localizar = True
        AtribuiValor
    End If
    rsTabelaProventoDesconto.Close
    Set rsTabelaProventoDesconto = Nothing
End Function

