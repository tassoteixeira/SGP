VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDependente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mCodigoFuncionario As Integer
Private mCodigoDependente As Integer
Private mNome As String
Private mDataNascimento As Date
Private mGrauDependencia As Integer
Private mInvalido As Boolean

Private rsDependente As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get Empresa() As Integer
    Empresa = mEmpresa
End Property
Public Property Let Empresa(ByVal Valor As Integer)
    mEmpresa = Valor
End Property


Public Property Get CodigoFuncionario() As Integer
    CodigoFuncionario = mCodigoFuncionario
End Property
Public Property Let CodigoFuncionario(ByVal Valor As Integer)
    mCodigoFuncionario = Valor
End Property


Public Property Get CodigoDependente() As Integer
    CodigoDependente = mCodigoDependente
End Property
Public Property Let CodigoDependente(ByVal Valor As Integer)
    mCodigoDependente = Valor
End Property


Public Property Get Nome() As String
    Nome = mNome
End Property
Public Property Let Nome(ByVal Valor As String)
    mNome = Valor
End Property


Public Property Get DataNascimento() As Date
    DataNascimento = mDataNascimento
End Property
Public Property Let DataNascimento(ByVal Valor As Date)
    mDataNascimento = Valor
End Property


Public Property Get GrauDependencia() As Integer
    GrauDependencia = mGrauDependencia
End Property
Public Property Let GrauDependencia(ByVal Valor As Integer)
    mGrauDependencia = Valor
End Property


Public Property Get Invalido() As Boolean
    Invalido = mInvalido
End Property
Public Property Let Invalido(ByVal Valor As Boolean)
    mInvalido = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal nEmpresa As Integer, ByVal nCodigoFuncionario As Integer, ByVal nCodigoDependente As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & nEmpresa
    xCondicao = xCondicao & " AND [Codigo do Funcionario] = " & nCodigoFuncionario
    xCondicao = xCondicao & " AND [Codigo do Dependente] = " & nCodigoDependente
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
    xCondicao = xCondicao & " AND [Codigo do Funcionario] = " & mCodigoFuncionario
    xCondicao = xCondicao & " AND [Codigo do Dependente] < " & mCodigoDependente
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa DESC, [Codigo do Funcionario] DESC, [Codigo do Dependente] DESC")
    LocalizarAnterior = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarNome(ByVal nEmpresa As Integer, ByVal nCodigoFuncionario As Integer, ByVal nNome As String) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro
    
    LocalizarNome = False
    xCondicao = " WHERE Empresa = " & nEmpresa
    xCondicao = xCondicao & " AND [Codigo do Funcionario] = " & nCodigoFuncionario
    xCondicao = xCondicao & " AND Nome = " & preparaTexto(nNome)
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarNome = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND [Codigo do Funcionario] = " & mCodigoFuncionario
    xCondicao = xCondicao & " AND [Codigo do Dependente] > 0"
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa ASC, [Codigo do Funcionario] ASC, [Codigo do Dependente] ASC")
    LocalizarPrimeiro = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimoDependente(ByVal nEmpresa As Integer, ByVal nCodigoFuncionario As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarUltimoDependente = False
    xCondicao = " WHERE Empresa = " & nEmpresa
    xCondicao = xCondicao & " AND [Codigo do Funcionario] = " & nCodigoFuncionario
    xCondicao = xCondicao & " AND [Codigo do Dependente] < " & 9999
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa DESC, [Codigo do Funcionario] DESC, [Codigo do Dependente] DESC")
    LocalizarUltimoDependente = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarProximo() As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarProximo = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND [Codigo do Funcionario] = " & mCodigoFuncionario
    xCondicao = xCondicao & " AND [Codigo do Dependente] > " & mCodigoDependente
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa ASC, [Codigo do Funcionario] ASC, [Codigo do Dependente] ASC")
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
    xCondicao = xCondicao & " AND [Codigo do Funcionario] < " & 9999
    xCondicao = xCondicao & " AND [Codigo do Dependente] < " & 9999
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa DESC, [Codigo do Funcionario] DESC, [Codigo do Dependente] DESC")
    LocalizarUltimo = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Dependente ( Empresa, [Codigo do Funcionario], [Codigo do Dependente], Nome, [Data de Nascimento], "
    gSQL = gSQL & "[Grau de Dependencia], Invalido ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlNumero(1, mCodigoFuncionario, ", ")
    Call sqlNumero(1, mCodigoDependente, ", ")
    Call sqlTexto(1, mNome, ", ")
    Call sqlData(1, mDataNascimento, ", ")
    Call sqlNumero(1, mGrauDependencia, ", ")
    Call sqlBoolean(1, mInvalido, " )")
    gConn.Execute gSQL, Incluir, adCmdText + adExecuteNoRecords
    Incluir = True
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal nEmpresa As Integer, ByVal nCodigoFuncionario As Integer, ByVal nCodigoDependente As Integer) As Boolean

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Dependente SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlNumero(2, ", [Codigo do Funcionario] = ", mCodigoFuncionario)
    Call sqlNumero(2, ", [Codigo do Dependente] = ", mCodigoDependente)
    Call sqlTexto(2, ", Nome = ", mNome)
    Call sqlData(2, ", [Data de Nascimento] = ", mDataNascimento)
    Call sqlNumero(2, ", [Grau de Dependencia] = ", mGrauDependencia)
    Call sqlBoolean(2, ", Invalido = ", mInvalido)
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND [Codigo do Funcionario] = " & nCodigoFuncionario
    gSQL = gSQL & " AND [Codigo do Dependente] = " & nCodigoDependente
    gConn.Execute gSQL, Alterar, adCmdText + adExecuteNoRecords
    Alterar = True
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal nEmpresa As Integer, ByVal nCodigoFuncionario As Integer, ByVal nCodigoDependente As Integer) As Boolean

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM Dependente"
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND [Codigo do Funcionario] = " & nCodigoFuncionario
    gSQL = gSQL & " AND [Codigo do Dependente] = " & nCodigoDependente
    gConn.Execute gSQL, Excluir, adCmdText + adExecuteNoRecords
    Excluir = True
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, [Codigo do Funcionario], [Codigo do Dependente], Nome, [Data de Nascimento], [Grau de Dependencia], Invalido FROM Dependente" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsDependente = New adodb.Recordset
    rsDependente.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsDependente.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsDependente.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsDependente("Empresa").Value
    CodigoFuncionario = rsDependente("Codigo do Funcionario").Value
    CodigoDependente = rsDependente("Codigo do Dependente").Value
    Nome = rsDependente("Nome").Value
    DataNascimento = rsDependente("Data de Nascimento").Value
    GrauDependencia = rsDependente("Grau de Dependencia").Value
    Invalido = rsDependente("Invalido").Value
End Sub


Private Function Localizar() As Boolean
    Localizar = False
    If AtualizaRecordset(1) Then
        Localizar = True
        AtribuiValor
    End If
    rsDependente.Close
    Set rsDependente = Nothing
End Function

