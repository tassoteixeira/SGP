VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEncerranteAtual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mCodigoBomba As Integer
Private mEncerrante As Currency
Private mSituacao As String
Private mLitragem As Currency

Private rsEncerranteAtual As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get Empresa() As Integer
    Empresa = mEmpresa
End Property
Public Property Let Empresa(ByVal Valor As Integer)
    mEmpresa = Valor
End Property


Public Property Get CodigoBomba() As Integer
    CodigoBomba = mCodigoBomba
End Property
Public Property Let CodigoBomba(ByVal Valor As Integer)
    mCodigoBomba = Valor
End Property


Public Property Get Encerrante() As Currency
    Encerrante = mEncerrante
End Property
Public Property Let Encerrante(ByVal Valor As Currency)
    mEncerrante = Valor
End Property


Public Property Get Situacao() As String
    Situacao = mSituacao
End Property
Public Property Let Situacao(ByVal Valor As String)
    mSituacao = Valor
End Property


Public Property Get Litragem() As Currency
    Litragem = mLitragem
End Property
Public Property Let Litragem(ByVal Valor As Currency)
    mLitragem = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal pEmpresa As Integer, ByVal pCodigoBomba As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND [Codigo da Bomba] = " & pCodigoBomba
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo = True Then
        AtribuiValor
    End If
    rsEncerranteAtual.Close
    Set rsEncerranteAtual = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarAnterior() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarAnterior = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND [Codigo da Bomba] < " & mCodigoBomba
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Codigo da Bomba] DESC")
    LocalizarAnterior = Localizar(1)
    If LocalizarAnterior = True Then
        AtribuiValor
    End If
    rsEncerranteAtual.Close
    Set rsEncerranteAtual = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Codigo da Bomba] ASC")
    LocalizarPrimeiro = Localizar(1)
    If LocalizarPrimeiro = True Then
        AtribuiValor
    End If
    rsEncerranteAtual.Close
    Set rsEncerranteAtual = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarProximo() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarProximo = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND [Codigo da Bomba] > " & mCodigoBomba
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Codigo da Bomba] ASC")
    LocalizarProximo = Localizar(1)
    If LocalizarProximo = True Then
        AtribuiValor
    End If
    rsEncerranteAtual.Close
    Set rsEncerranteAtual = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal pEmpresa As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY [Codigo da Bomba] DESC")
    LocalizarUltimo = Localizar(1)
    If LocalizarUltimo = True Then
        AtribuiValor
    End If
    rsEncerranteAtual.Close
    Set rsEncerranteAtual = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Encerrante_Atual ( Empresa, [Codigo da Bomba], Encerrante, Situacao, Litragem ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlNumero(1, mCodigoBomba, ", ")
    Call sqlValor(1, mEncerrante, ", ")
    Call sqlTexto(1, mSituacao, ", ")
    Call sqlValor(1, mLitragem, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal pEmpresa As Integer, ByVal pCodigoBomba As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Encerrante_Atual SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlNumero(2, ", [Codigo da Bomba] = ", mCodigoBomba)
    Call sqlValor(2, ", Encerrante = ", mEncerrante)
    Call sqlTexto(2, ", Situacao = ", mSituacao)
    Call sqlValor(2, ", Litragem = ", mLitragem)
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Codigo da Bomba] = " & pCodigoBomba
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function AlterarSituacao(ByVal pEmpresa As Integer, ByVal pCodigoBomba As Integer, ByVal pSituacao As String, ByVal pLitragem As Currency) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_AlterarSituacao

    AlterarSituacao = False
    gSQL = "UPDATE Encerrante_Atual SET "
    Call sqlTexto(2, "Situacao = ", pSituacao)
    Call sqlValor(2, ", Litragem = ", pLitragem)
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Codigo da Bomba] = " & pCodigoBomba
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        AlterarSituacao = True
    Else
        Call CriaLogCadastroDll2("Classe:cEncerranteAtual - Método:AlterarSituacao", "Não foi alterado bico=" & pCodigoBomba, gSQL)
    End If
    Exit Function

Err_AlterarSituacao:
    Call CriaLogCadastroDll2("Classe:cEncerranteAtual - Método:AlterarSituacao", Err.Description, gSQL)
End Function


Public Function Excluir(ByVal pEmpresa As Integer, ByVal pCodigoBomba As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM Encerrante_Atual"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND [Codigo da Bomba] = " & pCodigoBomba
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, [Codigo da Bomba], Encerrante, Situacao, Litragem FROM Encerrante_Atual" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal pQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsEncerranteAtual = New adodb.Recordset
    rsEncerranteAtual.CursorLocation = adUseClient
    i = Len(gSQL)
    If pQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & pQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsEncerranteAtual.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsEncerranteAtual.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsEncerranteAtual("Empresa").Value
    CodigoBomba = rsEncerranteAtual("Codigo da Bomba").Value
    Encerrante = rsEncerranteAtual("Encerrante").Value
    Situacao = rsEncerranteAtual("Situacao").Value
    Litragem = rsEncerranteAtual("Litragem").Value
End Sub


Private Function Localizar(ByVal pQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(pQtdRegistro) Then
        Localizar = True
    End If
End Function

