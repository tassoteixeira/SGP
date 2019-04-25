VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cIntegracaoNuvem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mData As Date
Private mNomeTabela As String
Private mChaveAcesso As String
Private mTipoOperacao As String
Private mIntegradoEm As Date

Dim lRecordsAffected As Long

Private rsIntegracaoNuvem As New adodb.Recordset


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


Public Property Get NomeTabela() As String
    NomeTabela = mNomeTabela
End Property
Public Property Let NomeTabela(ByVal Valor As String)
    mNomeTabela = Valor
End Property


Public Property Get ChaveAcesso() As String
    ChaveAcesso = mChaveAcesso
End Property
Public Property Let ChaveAcesso(ByVal Valor As String)
    mChaveAcesso = Valor
End Property


Public Property Get TipoOperacao() As String
    TipoOperacao = mTipoOperacao
End Property
Public Property Let TipoOperacao(ByVal Valor As String)
    mTipoOperacao = Valor
End Property


Public Property Get IntegradoEm() As Date
    IntegradoEm = mIntegradoEm
End Property
Public Property Let IntegradoEm(ByVal Valor As Date)
    mIntegradoEm = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pNomeTabela As String, ByVal pChaveAcesso As String, ByVal pTipoOperacao As String) As Boolean
    Dim xCondicao As String
    
    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & pEmpresa
    xCondicao = xCondicao & " AND Data = " & preparaData(pData)
    xCondicao = xCondicao & " [Nome da Tabela] = " & pNomeTabela
    xCondicao = xCondicao & " [Chave de Acesso] = " & pChaveAcesso
    xCondicao = xCondicao & " [Tipo de Operacao] = " & pTipoOperacao
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
End Function


Public Function Incluir() As Boolean

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO IntegracaoNuvem ( Empresa, Data, [Nome da Tabela], [Chave de Acesso], [Tipo de Operacao], "
    gSQL = gSQL & "[Integrado Em] ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlDataHora(1, mData, ", ")
    Call sqlTexto(1, mNomeTabela, ", ")
    Call sqlTexto(1, mChaveAcesso, ", ")
    Call sqlTexto(1, mTipoOperacao, ", ")
    Call sqlDataHora(1, mIntegradoEm, " )")
    gConnNuvem.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    Else
        Call CriaLogCadastroDll2("IntegracaoNuvem:Incluir - Erro ao incluir registro.", Err.Description, gSQL)
    End If
    Exit Function

Err_Incluir:
    Call CriaLogCadastroDll2("IntegracaoNuvem:Incluir - Erro não identificado.", Err.Description, gSQL)
End Function


Public Function Alterar(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pNomeTabela As String, ByVal pChaveAcesso As String, ByVal pTipoOperacao As String) As Boolean

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE IntegracaoNuvem SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlData(2, ", Data = ", mData)
    Call sqlTexto(2, ", [Nome da Tabela] = ", mNomeTabela)
    Call sqlTexto(2, ", [Chave de Acesso] = ", mChaveAcesso)
    Call sqlTexto(2, ", [Tipo de Operacao] = ", mTipoOperacao)
    Call sqlData(2, ", [Integrado Em] = ", mIntegradoEm)
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND [Nome da Tabela] = " & preparaTexto(pNomeTabela)
    gSQL = gSQL & " AND [Chave de Acesso] = " & preparaTexto(pChaveAcesso)
    gSQL = gSQL & " AND [Tipo de Operacao] = " & preparaTexto(pTipoOperacao)
    gConnNuvem.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    Else
        Call CriaLogCadastroDll2("IntegracaoNuvem:Alterar - Erro ao alterar registro.", Err.Description, gSQL)
    End If
    Exit Function

Err_Alterar:
    Call CriaLogCadastroDll2("IntegracaoNuvem:Alterar - Erro não identificado.", Err.Description, gSQL)
End Function


Public Function Excluir(ByVal pEmpresa As Integer, ByVal pData As Date, ByVal pNomeTabela As String, ByVal pChaveAcesso As String, ByVal pTipoOperacao As String) As Boolean

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM IntegracaoNuvem"
    gSQL = gSQL & " WHERE Empresa = " & pEmpresa
    gSQL = gSQL & " AND Data = " & preparaData(pData)
    gSQL = gSQL & " AND [Nome da Tabela] = " & preparaTexto(pNomeTabela)
    gSQL = gSQL & " AND [Chave de Acesso] = " & preparaTexto(pChaveAcesso)
    gSQL = gSQL & " AND [Tipo de Operacao] = " & preparaTexto(pTipoOperacao)
    gConnNuvem.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    Else
        Call CriaLogCadastroDll2("IntegracaoNuvem:Excluir - Erro ao excluir registro.", Err.Description, gSQL)
    End If
    Exit Function

Err_Excluir:
    Call CriaLogCadastroDll2("IntegracaoNuvem:Excluir - Erro não identificado.", Err.Description, gSQL)
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, Data, [Nome da Tabela], [Chave de Acesso], [Tipo de Operacao], [Integrado Em] FROM IntegracaoNuvem" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal pQtdRegistro As Integer) As Boolean
    Dim i As Integer

On Error GoTo Err_AtualizaRecordset
    
    AtualizaRecordset = False
    Set rsIntegracaoNuvem = New adodb.Recordset
    rsIntegracaoNuvem.CursorLocation = adUseClient
    i = Len(gSQL)
    If pQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & pQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsIntegracaoNuvem.Open gSQL, gConnNuvem, adOpenForwardOnly, adLockReadOnly
    If Not rsIntegracaoNuvem.EOF Then
        AtualizaRecordset = True
    End If
    Exit Function
    
Err_AtualizaRecordset:
    Call CriaLogCadastroDll2("IntegracaoNuvem:AtualizaRecordset - Erro não identificado.", Err.Description, gSQL)
End Function


Private Sub AtribuiValor()
    Empresa = rsIntegracaoNuvem("Empresa").Value
    Data = rsIntegracaoNuvem("Data").Value
    NomeTabela = rsIntegracaoNuvem("Nome da Tabela").Value
    ChaveAcesso = rsIntegracaoNuvem("Chave de Acesso").Value
    TipoOperacao = rsIntegracaoNuvem("Tipo de Operacao").Value
    If IsNull(rsIntegracaoNuvem("Integrado Em").Value) Then
        IntegradoEm = "00:00:00"
    Else
        IntegradoEm = rsIntegracaoNuvem("Integrado Em").Value
    End If
End Sub


Private Function Localizar(ByVal pQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(pQtdRegistro) Then
        Localizar = True
        AtribuiValor
    End If
    rsIntegracaoNuvem.Close
    Set rsIntegracaoNuvem = Nothing
End Function

