VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPlanoConta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mCodigo As String
Private mNome As String
Private mCodigoGrupo As String
Private mTipoConta As String
Private mContaReduzida As Integer

Private rsPlanoConta As New adodb.Recordset
'Const errNao_Consegui_Localizar = vbObjectError + 512 + 2
'Const errNao_Consegui_Salvar_Dados = vbObjectError + 512 + 3



'Inicio Propriedades da Classe
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


Public Property Get CodigoGrupo() As String
    CodigoGrupo = mCodigoGrupo
End Property
Public Property Let CodigoGrupo(ByVal Valor As String)
    mCodigoGrupo = Valor
End Property


Public Property Get TipoConta() As String
    TipoConta = mTipoConta
End Property
Public Property Let TipoConta(ByVal Valor As String)
    mTipoConta = Valor
End Property


Public Property Get ContaReduzida() As Integer
    ContaReduzida = mContaReduzida
End Property
Public Property Let ContaReduzida(ByVal Valor As Integer)
    mContaReduzida = Valor
End Property
'Fim Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal nEmpresa As Integer, ByVal nCodigo As String) As Boolean
Dim xCondicao As String

On Error GoTo trata_erro
    
    LocalizarCodigo = False
    xCondicao = ""
    xCondicao = xCondicao & "WHERE Empresa = " & nEmpresa
    xCondicao = xCondicao & "  AND Codigo = " & preparaTexto(nCodigo)
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarContaReduzida(ByVal nEmpresa As Integer, ByVal nContaReduzida As Integer) As Boolean
Dim xCondicao As String

On Error GoTo trata_erro
    
    LocalizarContaReduzida = False
    xCondicao = ""
    xCondicao = xCondicao & "WHERE Empresa = " & nEmpresa
    xCondicao = xCondicao & "  AND [Conta Reduzida] = " & nContaReduzida
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarContaReduzida = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarAnterior() As Boolean
Dim xCondicao As String

On Error GoTo trata_erro
    
    LocalizarAnterior = False
    xCondicao = ""
    xCondicao = xCondicao & "WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & "  AND Codigo < " & preparaTexto(mCodigo)
    gSQL = PreparaSQL(xCondicao, "ORDER BY Codigo DESC")
    LocalizarAnterior = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarNome(ByVal nNome As String) As Boolean
Dim xCondicao As String

On Error GoTo trata_erro
    
    LocalizarNome = False
    xCondicao = ""
    xCondicao = xCondicao & "WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & "  AND Nome = " & preparaTexto(nNome)
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
    xCondicao = ""
    xCondicao = xCondicao & "WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & "  AND Codigo > " & preparaTexto(" ")
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
    xCondicao = ""
    xCondicao = xCondicao & "WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & "  AND Codigo > " & preparaTexto(mCodigo)
    gSQL = PreparaSQL(xCondicao, "ORDER BY Codigo ASC")
    LocalizarProximo = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal nEmpresa As Integer) As Boolean
Dim xCondicao As String

On Error GoTo trata_erro
    
    LocalizarUltimo = False
    xCondicao = ""
    xCondicao = xCondicao & "WHERE Empresa = " & nEmpresa
    xCondicao = xCondicao & "  AND Codigo < " & preparaTexto("ZZZZZZZZZ")
    gSQL = PreparaSQL(xCondicao, "ORDER BY Codigo DESC")
    LocalizarUltimo = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Plano_Conta ( Empresa, Codigo, Nome, [Codigo do Grupo], [Tipo da Conta], [Conta Reduzida] ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlTexto(1, mCodigo, ", ")
    Call sqlTexto(1, mNome, ", ")
    Call sqlTexto(1, mCodigoGrupo, ", ")
    Call sqlTexto(1, mTipoConta, ", ")
    Call sqlNumero(1, mContaReduzida, " )")
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
    gSQL = "UPDATE Plano_Conta SET "
    Call sqlTexto(2, "Codigo = ", mCodigo)
    Call sqlTexto(2, ", Nome = ", mNome)
    Call sqlTexto(2, ", [Codigo do Grupo] = ", mCodigoGrupo)
    Call sqlTexto(2, ", [Tipo da Conta] = ", mTipoConta)
    Call sqlNumero(2, ", [Conta Reduzida] = ", mContaReduzida)
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & "   AND Codigo = " & preparaTexto(nCodigo)
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
    gSQL = ""
    gSQL = gSQL & "DELETE"
    gSQL = gSQL & "  FROM Plano_Conta"
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & "   AND Codigo = " & preparaTexto(nCodigo)
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function



'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, Codigo, Nome, [Codigo do Grupo], [Tipo da Conta], [Conta Reduzida] FROM Plano_Conta" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsPlanoConta = New adodb.Recordset
    rsPlanoConta.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsPlanoConta.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsPlanoConta.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsPlanoConta("Empresa").Value
    Codigo = rsPlanoConta("Codigo").Value
    Nome = rsPlanoConta("Nome").Value
    CodigoGrupo = rsPlanoConta("Codigo do Grupo").Value
    TipoConta = rsPlanoConta("Tipo da Conta").Value
    ContaReduzida = rsPlanoConta("Conta Reduzida").Value
End Sub


Private Function Localizar() As Boolean
    Localizar = False
    If AtualizaRecordset(1) Then
        Localizar = True
        AtribuiValor
    End If
    rsPlanoConta.Close
    Set rsPlanoConta = Nothing
End Function

