VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPaginaLMC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mTipoCombustivel As String
Private mData As Date
Private mPagina As Long

Private rsPaginaLMC As New adodb.Recordset
Dim lRecordsAffected As Long


'Inicio das Propriedades da Classe
Public Property Get Empresa() As Integer
    Empresa = mEmpresa
End Property
Public Property Let Empresa(ByVal Valor As Integer)
    mEmpresa = Valor
End Property


Public Property Get TipoCombustivel() As String
    TipoCombustivel = mTipoCombustivel
End Property
Public Property Let TipoCombustivel(ByVal Valor As String)
    mTipoCombustivel = Valor
End Property


Public Property Get Data() As Date
    Data = mData
End Property
Public Property Let Data(ByVal Valor As Date)
    mData = Valor
End Property


Public Property Get Pagina() As Long
    Pagina = mPagina
End Property
Public Property Let Pagina(ByVal Valor As Long)
    mPagina = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal pEmpresa As Integer, ByVal pTipoCombustivel As String, ByVal pData As Date) As Boolean
Dim xCondicao As String

On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE empresa = " & pEmpresa
    xCondicao = xCondicao & " AND tipo_combustivel = " & preparaTexto(pTipoCombustivel)
    xCondicao = xCondicao & " AND data = " & preparaData(pData)
    gSQL = PreparaSQL(xCondicao, "ORDER BY data DESC")
    If Localizar(1) Then
        LocalizarCodigo = True
        AtribuiValor
    End If
    rsPaginaLMC.Close
    Set rsPaginaLMC = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarAntesData(ByVal pEmpresa As Integer, ByVal pTipoCombustivel As String, ByVal pData As Date) As Boolean
Dim xCondicao As String

On Error GoTo trata_erro

    LocalizarAntesData = False
    xCondicao = " WHERE empresa = " & pEmpresa
    xCondicao = xCondicao & " AND tipo_combustivel = " & preparaTexto(pTipoCombustivel)
    xCondicao = xCondicao & " AND data < " & preparaData(pData)
    gSQL = PreparaSQL(xCondicao, "ORDER BY data DESC")
    If Localizar(1) Then
        LocalizarAntesData = True
        AtribuiValor
    End If
    rsPaginaLMC.Close
    Set rsPaginaLMC = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO pagina_lmc ( empresa, tipo_combustivel, data, pagina ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlTexto(1, mTipoCombustivel, ", ")
    Call sqlData(1, mData, ", ")
    Call sqlNumero(1, mPagina, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal pEmpresa As Integer, ByVal pTipoCombustivel As String, ByVal pData As Date) As Boolean

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE pagina_lmc SET "
    Call sqlNumero(2, "empresa = ", mEmpresa)
    Call sqlTexto(2, ", tipo_combustivel = ", mTipoCombustivel)
    Call sqlData(2, ", data = ", mData)
    Call sqlNumero(2, ", pagina = ", mPagina)
    gSQL = gSQL & " WHERE empresa = " & pEmpresa
    gSQL = gSQL & " AND tipo_combustivel = " & preparaTexto(pTipoCombustivel)
    gSQL = gSQL & " AND data = " & preparaData(pData)
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal pEmpresa As Integer, ByVal pTipoCombustivel As String, ByVal pData As Date) As Boolean

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM pagina_lmc"
    gSQL = gSQL & " WHERE empresa = " & pEmpresa
    gSQL = gSQL & " AND tipo_combustivel = " & preparaTexto(pTipoCombustivel)
    gSQL = gSQL & " AND data = " & preparaData(pData)
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT empresa, tipo_combustivel, data, pagina FROM pagina_lmc" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal pQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsPaginaLMC = New adodb.Recordset
    rsPaginaLMC.CursorLocation = adUseClient
    i = Len(gSQL)
    If pQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & pQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsPaginaLMC.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsPaginaLMC.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsPaginaLMC("empresa").Value
    TipoCombustivel = rsPaginaLMC("tipo_combustivel").Value
    Data = rsPaginaLMC("data").Value
    Pagina = rsPaginaLMC("pagina").Value
End Sub


Private Function Localizar(ByVal pQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(pQtdRegistro) Then
        Localizar = True
    End If
End Function

