VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCodigo As Integer
Private mNome As String
Private mHoraEntrada As Date
Private mHoraSaida As Date
Private mSituacao As String
Private mSenha As String
Private mTipoAcesso As Integer
Private mDataAlteracaoSenha As Date


Dim lRecordsAffected As Long

Private rsUsuario As New adodb.Recordset


'Inicio Propriedades da Classe
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


Public Property Get HoraEntrada() As Date
    HoraEntrada = mHoraEntrada
End Property
Public Property Let HoraEntrada(ByVal Valor As Date)
    mHoraEntrada = Valor
End Property


Public Property Get HoraSaida() As Date
    HoraSaida = mHoraSaida
End Property
Public Property Let HoraSaida(ByVal Valor As Date)
    mHoraSaida = Valor
End Property


Public Property Get Situacao() As String
    Situacao = mSituacao
End Property
Public Property Let Situacao(ByVal Valor As String)
    mSituacao = Valor
End Property


Public Property Get Senha() As String
    Senha = mSenha
End Property
Public Property Let Senha(ByVal Valor As String)
    mSenha = Valor
End Property


Public Property Get TipoAcesso() As Integer
    TipoAcesso = mTipoAcesso
End Property
Public Property Let TipoAcesso(ByVal Valor As Integer)
    mTipoAcesso = Valor
End Property


Public Property Get DataAlteracaoSenha() As Date
    DataAlteracaoSenha = mDataAlteracaoSenha
End Property
Public Property Let DataAlteracaoSenha(ByVal Valor As Date)
    mDataAlteracaoSenha = Valor
End Property
'Fim Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal pCodigo As Integer) As Boolean
    LocalizarCodigo = False
    gSQL = PreparaSQL("WHERE Codigo = " & pCodigo, "")
    LocalizarCodigo = Localizar
End Function


Public Function LocalizarAnterior() As Boolean
    LocalizarAnterior = False
    gSQL = PreparaSQL("WHERE Codigo < " & mCodigo, "ORDER BY Codigo DESC")
    LocalizarAnterior = Localizar
End Function


Public Function LocalizarNome(ByVal pNome As String) As Boolean
    LocalizarNome = False
    gSQL = PreparaSQL("WHERE Nome = " & preparaTexto(pNome), "")
    LocalizarNome = Localizar
End Function


Public Function LocalizarPrimeiro() As Boolean
    LocalizarPrimeiro = False
    gSQL = PreparaSQL("WHERE Codigo > " & 0, "ORDER BY Codigo ASC")
    LocalizarPrimeiro = Localizar
End Function


Public Function LocalizarProximo() As Boolean
    LocalizarProximo = False
    gSQL = PreparaSQL("WHERE Codigo > " & mCodigo, "ORDER BY Codigo ASC")
    LocalizarProximo = Localizar
End Function


Public Function LocalizarUltimo() As Boolean
    LocalizarUltimo = False
    gSQL = PreparaSQL("WHERE Codigo < " & 9999, "ORDER BY Codigo DESC")
    LocalizarUltimo = Localizar
End Function


Public Function Incluir() As Boolean

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Usuario ( Codigo, Nome, [Hora de Entrada], [Hora de Saida], Situacao, Senha, [Tipo de Acesso], "
    gSQL = gSQL & "[Data da Alteracao da Senha] ) VALUES ( "
    Call sqlNumero(1, mCodigo, ", ")
    Call sqlTexto(1, mNome, ", ")
    Call sqlHora(1, mHoraEntrada, ", ")
    Call sqlHora(1, mHoraSaida, ", ")
    Call sqlTexto(1, mSituacao, ", ")
    Call sqlTexto(1, mSenha, ", ")
    Call sqlNumero(1, mTipoAcesso, ", ")
    Call sqlData(1, mDataAlteracaoSenha, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    Else
        Call CriaLogCadastroDll2("Usuario:Incluir - Erro ao incluir registro.", Err.Description, gSQL)
    End If
    Exit Function

Err_Incluir:
    Call CriaLogCadastroDll2("Usuario:Incluir - Erro não identificado.", Err.Description, gSQL)
End Function


Public Function Alterar(ByVal pCodigo As Integer) As Boolean

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Usuario SET "
    Call sqlTexto(2, "Nome = ", mNome)
    Call sqlHora(2, ", [Hora de Entrada] = ", mHoraEntrada)
    Call sqlHora(2, ", [Hora de Saida] = ", mHoraSaida)
    Call sqlTexto(2, ", Situacao = ", mSituacao)
    Call sqlTexto(2, ", Senha = ", mSenha)
    Call sqlNumero(2, ", [Tipo de Acesso] = ", mTipoAcesso)
    Call sqlData(2, ", [Data da Alteracao da Senha] = ", mDataAlteracaoSenha)
    gSQL = gSQL & " WHERE Codigo = " & pCodigo
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    Else
        Call CriaLogCadastroDll2("Usuario:Alterar - Erro ao alterar registro.", Err.Description, gSQL)
    End If
    Exit Function

Err_Alterar:
    Call CriaLogCadastroDll2("Usuario:Alterar - Erro não identificado.", Err.Description, gSQL)
End Function


Public Function Excluir(ByVal pCodigo As Integer) As Boolean

On Error GoTo Err_Excluir
            
    Excluir = False
    gSQL = "DELETE FROM Usuario WHERE Codigo = " & pCodigo
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    Else
        Call CriaLogCadastroDll2("Usuario:Excluir - Erro ao excluir registro.", Err.Description, gSQL)
    End If
    Exit Function

Err_Excluir:
    Call CriaLogCadastroDll2("Usuario:Excluir - Erro não identificado.", Err.Description, gSQL)
End Function



'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Codigo, Nome, [Hora de Entrada], [Hora de Saida], Situacao, Senha, [Tipo de Acesso], [Data da Alteracao da Senha] FROM Usuario" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal pQtdRegistro As Integer) As Boolean
    Dim i As Integer

On Error GoTo Err_AtualizaRecordset
    
    AtualizaRecordset = False
    Set rsUsuario = New adodb.Recordset
    rsUsuario.CursorLocation = adUseClient
    i = Len(gSQL)
    If pQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & pQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsUsuario.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsUsuario.EOF Then
        AtualizaRecordset = True
    End If
    Exit Function
    
Err_AtualizaRecordset:
    Call CriaLogCadastroDll2("Usuario:AtualizaRecordset - Erro não identificado.", Err.Description, gSQL)
End Function


Private Sub AtribuiValor()
    Dim xLocal As Integer
    
On Error GoTo Err_AtribuiValor
    
    xLocal = 1
    Codigo = rsUsuario("Codigo").Value
    xLocal = 2
    Nome = rsUsuario("Nome").Value
    xLocal = 3
    HoraEntrada = rsUsuario("Hora de Entrada").Value
    xLocal = 4
    HoraSaida = rsUsuario("Hora de Saida").Value
    xLocal = 5
    Situacao = rsUsuario("Situacao").Value
    xLocal = 6
    Senha = rsUsuario("Senha").Value
    xLocal = 7
    TipoAcesso = rsUsuario("Tipo de Acesso").Value
    xLocal = 8
    DataAlteracaoSenha = rsUsuario("Data da Alteracao da Senha").Value
    Exit Sub
    
Err_AtribuiValor:
    Call CriaLogCadastroDll2("Usuario:AtribuiValor - Erro não identificado.", Err.Description, "Local:" & xLocal)
End Sub


Private Function Localizar() As Boolean
    Localizar = False
    If AtualizaRecordset(1) Then
        Localizar = True
        AtribuiValor
    End If
    rsUsuario.Close
    Set rsUsuario = Nothing
End Function

