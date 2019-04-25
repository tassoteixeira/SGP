VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTipoDocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCodigo As Integer
Private mNome As String

Private rsTipoDocumento As New adodb.Recordset
'Const errNao_Consegui_Localizar = vbObjectError + 512 + 2
'Const errNao_Consegui_Salvar_Dados = vbObjectError + 512 + 3



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
'Fim Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal nCodigo As Integer) As Boolean

On Error GoTo trata_erro
    
    LocalizarCodigo = False
    gSQL = PreparaSQL("WHERE Codigo = " & nCodigo, "")
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


Public Function LocalizarNome(ByVal nNome As String) As Boolean

On Error GoTo trata_erro
    
    LocalizarNome = False
    gSQL = PreparaSQL("WHERE Nome = " & preparaTexto(nNome), "")
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

On Error GoTo Err_Incluir

Dim lRecordsAffected As Long

    Incluir = False
    gSQL = "INSERT INTO Tipo_Documentos ( Codigo, Nome ) VALUES ( "
    Call sqlNumero(1, mCodigo, ", ")
    Call sqlTexto(1, mNome, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    Incluir = True
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal nCodigo As Integer) As Boolean

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Tipo_Documentos SET "
    Call sqlTexto(2, "Nome = ", mNome)
    gSQL = gSQL & " WHERE Codigo = " & nCodigo
    gConn.Execute gSQL, Alterar, adCmdText + adExecuteNoRecords
    Alterar = True
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal nCodigo As Integer) As Boolean

On Error GoTo Err_Excluir
            
    Excluir = False
    gSQL = "DELETE FROM Tipo_Documentos WHERE Codigo = " & nCodigo
    gConn.Execute gSQL, Excluir, adCmdText + adExecuteNoRecords
    Excluir = True
    Exit Function

Err_Excluir:
End Function



'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Codigo, Nome FROM Tipo_Documentos" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsTipoDocumento = New adodb.Recordset
    rsTipoDocumento.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsTipoDocumento.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsTipoDocumento.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Codigo = rsTipoDocumento("Codigo").Value
    Nome = rsTipoDocumento("Nome").Value
End Sub


Private Function Localizar() As Boolean
    Localizar = False
    If AtualizaRecordset(1) Then
        Localizar = True
        AtribuiValor
    End If
    rsTipoDocumento.Close
    Set rsTipoDocumento = Nothing
End Function

