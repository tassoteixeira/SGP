VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTributacaoPisCofins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim mIdEstabelecimento As Integer
Dim mCstCofins As String
Dim mCstPis As String
Dim mAliquotaCofins As Currency
Dim mAliquotaPis As Currency
Dim mEntrada As Boolean
Dim mSaida As Boolean


Private rsTributacaoPisCofins As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get IdEstabelecimento() As Integer
    IdEstabelecimento = mIdEstabelecimento
End Property
Public Property Let IdEstabelecimento(ByVal Valor As Integer)
    mIdEstabelecimento = Valor
End Property
Public Property Get CstCofins() As String
    CstCofins = mCstCofins
End Property
Public Property Let CstCofins(ByVal Valor As String)
    mCstCofins = Valor
End Property
Public Property Get CstPis() As String
    CstPis = mCstPis
End Property
Public Property Let CstPis(ByVal Valor As String)
    mCstPis = Valor
End Property

Public Property Get AliquotaCofins() As Currency
    AliquotaCofins = mAliquotaCofins
End Property
Public Property Let AliquotaCofins(ByVal Valor As Currency)
    mAliquotaCofins = Valor
End Property

Public Property Get AliquotaPis() As Currency
    AliquotaPis = mAliquotaPis
End Property
Public Property Let AliquotaPis(ByVal Valor As Currency)
    mAliquotaPis = Valor
End Property

Public Property Get Entrada() As Boolean
    Entrada = mEntrada
End Property
Public Property Let Entrada(ByVal Valor As Boolean)
    mEntrada = Valor
End Property

Public Property Get Saida() As Boolean
    Saida = mSaida
End Property
Public Property Let Saida(ByVal Valor As Boolean)
    mSaida = Valor
End Property



Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    

    gSQL = "INSERT INTO TributacaoPisCofins (IdEstabelecimento_TributacaoPisCofins, CstCofins_TributacaoPisCofins, CstPis_TributacaoPisCofins, AliquotaCofins_TributacaoPisCofins, AliquotaPis_TributacaoPisCofins, Entrada_TributacaoPisCofins, Saida_TributacaoPisCofins) "
    gSQL = gSQL & ") VALUES ("
    Call sqlNumero(1, mIdEstabelecimento, ", ")
    Call sqlTexto(1, mCstCofins, ", ")
    Call sqlTexto(1, mCstPis, ", ")
    Call sqlValor(1, mAliquotaCofins, ", ")
    Call sqlValor(1, mAliquotaPis, ", ")
    Call sqlBoolean(1, mEntrada, ", ")
    Call sqlBoolean(1, mSaida, " )")
    
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    
    If lRecordsAffected > 0 Then
       Incluir = True
    End If
    Exit Function
    
Err_Incluir:
End Function

Public Function Alterar(ByVal pIdEstabelecimento As Integer, ByVal pCSTCofins As String, ByVal pCSTPis As String, ByVal pEntrada As Boolean, ByVal pSaida As Boolean) As Boolean
    Dim lRecordsAffected As Long
    
    On Error GoTo Err_Alterar
    
    Alterar = False
    
    gSQL = "UPDATE TributacaoPisCofins SET "
    Call sqlNumero(2, "IdEstabelecimento_TributacaoPisCofins = ", mIdEstabelecimento)
    Call sqlTexto(2, ",  CstCofins_TributacaoPisCofins = ", mCstCofins)
    Call sqlTexto(2, ",  CstPis_TributacaoPisCofins = ", mCstPis)
    Call sqlValor(2, ",  AliquotaCofins_TributacaoPisCofins = ", mAliquotaCofins)
    Call sqlValor(2, ",  AliquotaPis_TributacaoPisCofins = ", mAliquotaPis)
    Call sqlBoolean(2, ",  Entrada_TributacaoPisCofins = ", mEntrada)
    Call sqlBoolean(2, ",  Saida_TributacaoPisCofins = ", mSaida)
    
    gSQL = gSQL & " WHERE IdEstabelecimento_TributacaoPisCofins = " & pIdEstabelecimento
    gSQL = gSQL & " AND CstCofins_TributacaoPisCofins = " & preparaTexto(pCSTCofins)
    gSQL = gSQL & " AND CstPis_TributacaoPisCofins = " & preparaTexto(pCSTPis)
    gSQL = gSQL & " AND Entrada_TributacaoPisCofins = " & preparaBooleano(pEntrada)
    gSQL = gSQL & " AND Saida_TributacaoPisCofins = " & preparaBooleano(pSaida)
    
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function


Err_Alterar:

End Function

Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsTributacaoPisCofins = New adodb.Recordset
    rsTributacaoPisCofins.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsTributacaoPisCofins.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsTributacaoPisCofins.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()

    IdEstabelecimento = rsTributacaoPisCofins("IdEstabelecimento_TributacaoPisCofins").Value
    CstCofins = rsTributacaoPisCofins("CstCofins_TributacaoPisCofins").Value
    CstPis = rsTributacaoPisCofins("CstPis_TributacaoPisCofins").Value
    AliquotaCofins = rsTributacaoPisCofins("AliquotaCofins_TributacaoPisCofins").Value
    AliquotaPis = rsTributacaoPisCofins("AliquotaPis_TributacaoPisCofins").Value
    Entrada = rsTributacaoPisCofins("Entrada_TributacaoPisCofins").Value
    Saida = rsTributacaoPisCofins("Saida_TributacaoPisCofins").Value
End Sub
Public Function LocalizarCSTPis(ByVal pIdEstabelecimento As Integer, ByVal pCSTPis As String, ByVal pEntrada As Boolean, ByVal pSaida As Boolean) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCSTPis = False
   
    xCondicao = xCondicao & " WHERE IdEstabelecimento_TributacaoPisCofins = " & pIdEstabelecimento
    xCondicao = xCondicao & " AND CstPis_TributacaoPisCofins = " & preparaTexto(pCSTPis)
    xCondicao = xCondicao & " AND Entrada_TributacaoPisCofins = " & preparaBooleano(pEntrada)
    xCondicao = xCondicao & " AND Saida_TributacaoPisCofins = " & preparaBooleano(pSaida)

    
    xCondicao = PreparaSQL(xCondicao, "")
    LocalizarCSTPis = Localizar(1)
    If LocalizarCSTPis Then
        AtribuiValor
    End If
    rsTributacaoPisCofins.Close
    Set rsTributacaoPisCofins = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function
Public Function LocalizarCSTCofins(ByVal pIdEstabelecimento As Integer, ByVal pCSTCofins As String, ByVal pEntrada As Boolean, ByVal pSaida As Boolean) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCSTCofins = False
   
    xCondicao = xCondicao & " WHERE IdEstabelecimento_TributacaoPisCofins = " & pIdEstabelecimento
    xCondicao = xCondicao & " AND CstCofins_TributacaoPisCofins = " & preparaTexto(pCSTCofins)
    xCondicao = xCondicao & " AND Entrada_TributacaoPisCofins = " & preparaBooleano(pEntrada)
    xCondicao = xCondicao & " AND Saida_TributacaoPisCofins = " & preparaBooleano(pSaida)

    
    xCondicao = PreparaSQL(xCondicao, "")
    LocalizarCSTCofins = Localizar(1)
    If LocalizarCSTCofins Then
        AtribuiValor
    End If
    rsTributacaoPisCofins.Close
    Set rsTributacaoPisCofins = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function

Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    Dim xSQL As String
    
    
    xSQL = "SELECT IdEstabelecimento_TributacaoPisCofins, CstCofins_TributacaoPisCofins, CstPis_TributacaoPisCofins,AliquotaCofins_TributacaoPisCofins, AliquotaPis_TributacaoPisCofins, Entrada_TributacaoPisCofins, Saida_TributacaoPisCofins "
    xSQL = xSQL & " FROM TributacaoPisCofins"
    
    PreparaSQL = Trim(xSQL & " " & xCondicao & " " & xTipoOrdem)
    
End Function



