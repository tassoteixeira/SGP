VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCfop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCodigo As String
Private mNaturezaOperacaoReduzida As String
Private mNaturezaOperacao As String
Private mDescricaoNaturezaOperacao As String
Private mOperacao As String



Private rsCfop As New adodb.Recordset
'Const errNao_Consegui_Localizar = vbObjectError + 512 + 2
'Const errNao_Consegui_Salvar_Dados = vbObjectError + 512 + 3



'Inicio Propriedades da Classe
Public Property Get Codigo() As String
    Codigo = mCodigo
End Property
Public Property Let Codigo(ByVal Valor As String)
    mCodigo = Valor
End Property


Public Property Get NaturezaOperacaoReduzida() As String
    NaturezaOperacaoReduzida = mNaturezaOperacaoReduzida
End Property
Public Property Let NaturezaOperacaoReduzida(ByVal Valor As String)
    mNaturezaOperacaoReduzida = Valor
End Property


Public Property Get NaturezaOperacao() As String
    NaturezaOperacao = mNaturezaOperacao
End Property
Public Property Let NaturezaOperacao(ByVal Valor As String)
    mNaturezaOperacao = Valor
End Property


Public Property Get DescricaoNaturezaOperacao() As String
    DescricaoNaturezaOperacao = mDescricaoNaturezaOperacao
End Property
Public Property Let DescricaoNaturezaOperacao(ByVal Valor As String)
    mDescricaoNaturezaOperacao = Valor
End Property


Public Property Get Operacao() As String
    Operacao = mOperacao
End Property
Public Property Let Operacao(ByVal Valor As String)
    mOperacao = Valor
End Property


'Fim Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal pCodigo As String) As Boolean

On Error GoTo trata_erro
    
    LocalizarCodigo = False
    gSQL = PreparaSQL("WHERE Codigo = " & preparaTexto(pCodigo), "")
    LocalizarCodigo = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarAnterior() As Boolean

On Error GoTo trata_erro
    
    LocalizarAnterior = False
    gSQL = PreparaSQL("WHERE Codigo < " & mCodigo, "ORDER BY Codigo DESC")
    LocalizarAnterior = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarNome(ByVal pNome As String) As Boolean

On Error GoTo trata_erro
    
    LocalizarNome = False
    gSQL = PreparaSQL("WHERE [Natureza da Operacao Reduzida] = " & preparaTexto(pNome), "")
    LocalizarNome = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean

On Error GoTo trata_erro
    
    LocalizarPrimeiro = False
    gSQL = PreparaSQL("WHERE Codigo > " & 0, "ORDER BY Codigo ASC")
    LocalizarPrimeiro = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarProximo() As Boolean

On Error GoTo trata_erro
    
    LocalizarProximo = False
    gSQL = PreparaSQL("WHERE Codigo > " & mCodigo, "ORDER BY Codigo ASC")
    LocalizarProximo = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo() As Boolean

On Error GoTo trata_erro
    
    LocalizarUltimo = False
    gSQL = PreparaSQL("WHERE Codigo < " & 9999, "ORDER BY Codigo DESC")
    LocalizarUltimo = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Cfop ( Codigo, [Natureza da Operacao Reduzida], [Natureza da Operacao], "
    gSQL = gSQL & "[Descricao da Natureza da Operacao], Operacao ) VALUES ( "
    Call sqlTexto(1, mCodigo, ", ")
    Call sqlTexto(1, mNaturezaOperacaoReduzida, ", ")
    Call sqlTexto(1, mNaturezaOperacao, ", ")
    Call sqlTexto(1, mDescricaoNaturezaOperacao, ", ")
    Call sqlTexto(1, mOperacao, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Incluir = True
    End If
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal pCodigo As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Grupo SET "
    Call sqlTexto(2, "Codigo = ", mCodigo)
    Call sqlTexto(2, ", [NaturezaOperacaoReduzida] = ", mNaturezaOperacaoReduzida)
    Call sqlTexto(2, ", [NaturezaOperacao] = ", mNaturezaOperacao)
    Call sqlTexto(2, ", [DescricaoNaturezaOperacao] = ", mDescricaoNaturezaOperacao)
    Call sqlTexto(2, ", [Operacao] = ", mOperacao)
    gSQL = gSQL & " WHERE Codigo = " & pCodigo
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal nCodigo As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Excluir
            
    Excluir = False
    gSQL = "DELETE FROM Cfop WHERE Codigo = " & nCodigo
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function



'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Codigo, [Natureza da Operacao Reduzida], [Natureza da Operacao], [Descricao da Natureza da Operacao], Operacao FROM Cfop" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsCfop = New adodb.Recordset
    rsCfop.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsCfop.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsCfop.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Codigo = rsCfop("Codigo").Value
    NaturezaOperacaoReduzida = rsCfop("Natureza da Operacao Reduzida").Value
    NaturezaOperacao = rsCfop("Natureza da Operacao").Value
    DescricaoNaturezaOperacao = rsCfop("Descricao da Natureza da Operacao").Value
    Operacao = rsCfop("Operacao").Value
    End Sub


Private Function Localizar() As Boolean
    Localizar = False
    If AtualizaRecordset(1) Then
        Localizar = True
        AtribuiValor
    End If
    rsCfop.Close
    Set rsCfop = Nothing
End Function

