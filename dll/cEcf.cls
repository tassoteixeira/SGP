VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEcf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mCodigo As Integer
Private mNumeroSerie As String
Private mIlha As Integer
Private mIPPdv As String
Private rsECF As New adodb.Recordset

Private mModelo As String


'Inicio das Propriedades da Classe
Public Property Get Empresa() As Integer
    Empresa = mEmpresa
End Property
Public Property Let Empresa(ByVal Valor As Integer)
    mEmpresa = Valor
End Property


Public Property Get Codigo() As Integer
    Codigo = mCodigo
End Property
Public Property Let Codigo(ByVal Valor As Integer)
    mCodigo = Valor
End Property


Public Property Get NumeroSerie() As String
    NumeroSerie = mNumeroSerie
End Property
Public Property Let NumeroSerie(ByVal Valor As String)
    mNumeroSerie = Valor
End Property


Public Property Get Ilha() As Integer
    Ilha = mIlha
End Property
Public Property Let Ilha(ByVal Valor As Integer)
    mIlha = Valor
End Property


Public Property Get IPPdv() As String
    IPPdv = mIPPdv
End Property
Public Property Let IPPdv(ByVal Valor As String)
    mIPPdv = Valor
End Property


Public Property Get Modelo() As String
    Modelo = mModelo
End Property
Public Property Let Modelo(ByVal Valor As String)
    mModelo = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal pEmpresa As Integer, ByVal pCodigo As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND Codigo = " & pCodigo
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo = True Then
        AtribuiValor
    End If
    rsECF.Close
    Set rsECF = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarNumeroSerie(ByVal pEmpresa As Integer, ByVal pNumeroSerie As String) As Boolean
Dim xCondicao As String

On Error GoTo trata_erro
    
    LocalizarNumeroSerie = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND [Numero de Serie] = " & preparaTexto(pNumeroSerie)
    gSQL = PreparaSQL(xCondicao, "")
    If Localizar(1) Then
        LocalizarNumeroSerie = True
        AtribuiValor
    End If
    rsECF.Close
    Set rsECF = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarIpPdv(ByVal pEmpresa As Integer, ByVal pIpPdv As String) As Boolean
Dim xCondicao As String

On Error GoTo trata_erro
    
    LocalizarIpPdv = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND [IP do Pdv] = " & preparaTexto(pIpPdv)
    gSQL = PreparaSQL(xCondicao, "")
    If Localizar(1) Then
        LocalizarIpPdv = True
        AtribuiValor
    End If
    rsECF.Close
    Set rsECF = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal pEmpresa As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    gSQL = PreparaSQL(xCondicao, "ORDER BY Codigo DESC")
    LocalizarUltimo = Localizar(1)
    If LocalizarUltimo = True Then
        AtribuiValor
    End If
    rsECF.Close
    Set rsECF = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Ecf ( Empresa, Codigo, [Numero de Serie], Ilha, [IP do Pdv], "
    gSQL = gSQL & "Modelo ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlNumero(1, mCodigo, ", ")
    Call sqlTexto(1, mNumeroSerie, ", ")
    Call sqlNumero(1, mIlha, ", ")
    Call sqlTexto(1, mIPPdv, ", ")
    Call sqlTexto(1, mModelo, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal pEmpresa As Integer, ByVal pCodigo As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE ECF SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlNumero(2, ", Codigo = ", mCodigo)
    Call sqlTexto(2, ", [Numero de Serie] = ", mNumeroSerie)
    Call sqlNumero(2, ", Ilha = ", mIlha)
    Call sqlTexto(2, ", [IP do Pdv] = ", mIPPdv)
    Call sqlTexto(2, ", Modelo = ", mModelo)
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & "   AND Codigo = " & pCodigo
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal pEmpresa As Integer, ByVal pCodigo As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM ECF"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & "   AND Codigo = " & pCodigo
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, Codigo, [Numero de Serie], Ilha, [IP do Pdv], Modelo FROM Ecf" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal pQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsECF = New adodb.Recordset
    rsECF.CursorLocation = adUseClient
    i = Len(gSQL)
    If pQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & pQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsECF.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsECF.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsECF("Empresa").Value
    Codigo = rsECF("Codigo").Value
    NumeroSerie = rsECF("Numero de Serie").Value
    Ilha = rsECF("Ilha").Value
    IPPdv = rsECF("IP do Pdv").Value
    Modelo = rsECF("Modelo").Value
End Sub


Private Function Localizar(ByVal pQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(pQtdRegistro) Then
        Localizar = True
    End If
End Function

