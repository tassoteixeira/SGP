VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGradePrecoAutomacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCodigo As Integer
Private mTipoCombustivel As String
Private mNome As String
Private mNomeCurto As String
Private mPrecoVista As Currency
Private mPrecoPrazo As Currency


Private rsGradePrecoAutomacao As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get Codigo() As Integer
    Codigo = mCodigo
End Property
Public Property Let Codigo(ByVal Valor As Integer)
    mCodigo = Valor
End Property

Public Property Get TipoCombustivel() As String
    TipoCombustivel = mTipoCombustivel
End Property
Public Property Let TipoCombustivel(ByVal Valor As String)
    mTipoCombustivel = Valor
End Property

Public Property Get Nome() As String
    Nome = mNome
End Property
Public Property Let Nome(ByVal Valor As String)
    mNome = Valor
End Property

Public Property Get NomeCurto() As String
    NomeCurto = mNomeCurto
End Property
Public Property Let NomeCurto(ByVal Valor As String)
    mNomeCurto = Valor
End Property

Public Property Get PrecoVista() As Currency
    PrecoVista = mPrecoVista
End Property
Public Property Let PrecoVista(ByVal Valor As Currency)
    mPrecoVista = Valor
End Property

Public Property Get PrecoPrazo() As Currency
    PrecoPrazo = mPrecoPrazo
End Property
Public Property Let PrecoPrazo(ByVal Valor As Currency)
    mPrecoPrazo = Valor
End Property



'Inicio Métodos da Classe


Public Function Incluir() As Boolean
    
Dim lRecordsAffected As Long
On Error GoTo Err_Incluir

        Incluir = False
        gSQL = "INSERT INTO GradePrecoAutomacao(Codigo_GradePrecoAutomacao,TipoCombustivel_GradePrecoAutomacao,Nome_GradePrecoAutomacao, "
        gSQL = gSQL & "NomeCurto_GradePrecoAutomacao, PrecoVista_GradePrecoAutomacao, PrecoPrazo_GradePrecoAutomacao"
        gSQL = gSQL & ") VALUES ( "
        Call sqlNumero(1, mCodigo, ", ")
        Call sqlTexto(1, mTipoCombustivel, ", ")
        Call sqlTexto(1, mNome, ", ")
        Call sqlTexto(1, mNomeCurto, ", ")
        Call sqlValor(1, mPrecoVista, ", ")
        Call sqlValor(1, mPrecoPrazo, " )")

        gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords

        
        If lRecordsAffected > 0 Then
            Incluir = True
        End If
        Exit Function
        
Err_Incluir:
        

End Function
Public Function LocalizarCodigo(ByVal pCodigo As Integer) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Codigo_GradePrecoAutomacao = " & pCodigo
    
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo Then
        AtribuiValor
    End If
    rsGradePrecoAutomacao.Close
    Set rsGradePrecoAutomacao = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function
Public Function LocalizarTipoCombustivel(ByVal pTipoCombustivel As String) As Boolean
Dim xCondicao As String
On Error GoTo trata_erro

    LocalizarTipoCombustivel = False
    xCondicao = " WHERE TipoCombustivel_GradePrecoAutomacao = " & preparaTexto(pTipoCombustivel)
    
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarTipoCombustivel = Localizar(1)
    If LocalizarTipoCombustivel Then
        AtribuiValor
    End If
    rsGradePrecoAutomacao.Close
    Set rsGradePrecoAutomacao = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function
Public Function ExisteGradeCadastrada() As Boolean
Dim xRetornoConsulta As Boolean
On Error GoTo trata_erro

    ExisteGradeCadastrada = False
    xRetornoConsulta = False
    
    gSQL = "SELECT COUNT(1) AS QtdRegistro FROM GradePrecoAutomacao"
    xRetornoConsulta = Localizar(0)
    
    If xRetornoConsulta = True Then
        If rsGradePrecoAutomacao("QtdRegistro").Value > 0 Then
            ExisteGradeCadastrada = True
        End If
    End If
    
    rsGradePrecoAutomacao.Close
    Set rsGradePrecoAutomacao = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function

Public Function DefinirPrecoGrade(ByVal pCodigo As Integer, ByVal pPrecoVista As Currency, ByVal pPrecoPrazo As Currency) As Boolean

Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    DefinirPrecoGrade = False
    gSQL = "UPDATE GradePrecoAutomacao SET "
    Call sqlValor(2, "PrecoVista_GradePrecoAutomacao = ", pPrecoVista)
    Call sqlValor(2, ", PrecoPrazo_GradePrecoAutomacao = ", pPrecoPrazo)
    gSQL = gSQL & " WHERE Codigo_GradePrecoAutomacao = " & pCodigo
    
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        DefinirPrecoGrade = True
    End If
    Exit Function

Err_Alterar:
End Function


'Funções / Procedures internas
Private Sub AtribuiValor()

    Codigo = rsGradePrecoAutomacao("Codigo_GradePrecoAutomacao").Value
    TipoCombustivel = rsGradePrecoAutomacao("TipoCombustivel_GradePrecoAutomacao").Value
    Nome = rsGradePrecoAutomacao("Nome_GradePrecoAutomacao").Value
    NomeCurto = rsGradePrecoAutomacao("NomeCurto_GradePrecoAutomacao").Value
    If IsNull(rsGradePrecoAutomacao("PrecoVista_GradePrecoAutomacao").Value) Then
        PrecoVista = 0
    Else
        PrecoVista = rsGradePrecoAutomacao("PrecoVista_GradePrecoAutomacao").Value
    End If
    If IsNull(rsGradePrecoAutomacao("PrecoPrazo_GradePrecoAutomacao").Value) Then
        PrecoPrazo = 0
    Else
        PrecoPrazo = rsGradePrecoAutomacao("PrecoPrazo_GradePrecoAutomacao").Value
    End If
End Sub
Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function
Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsGradePrecoAutomacao = New adodb.Recordset
    rsGradePrecoAutomacao.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsGradePrecoAutomacao.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsGradePrecoAutomacao.EOF Then
        AtualizaRecordset = True
    End If
End Function

Private Function PreparaSQL(ByVal pCondicao As String, ByVal pTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Codigo_GradePrecoAutomacao,TipoCombustivel_GradePrecoAutomacao, Nome_GradePrecoAutomacao, NomeCurto_GradePrecoAutomacao, PrecoVista_GradePrecoAutomacao, PrecoPrazo_GradePrecoAutomacao FROM GradePrecoAutomacao" & " " & pCondicao & " " & pTipoOrdem)
End Function

