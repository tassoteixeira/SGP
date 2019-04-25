VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPercentualImposto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mTabela As Integer
Private mCodigo As String
Private mEx As Integer
Private mNome As String
Private mAliquotaNacional As Currency
Private mAliquotaImportada As Currency

Private rsPercentualImposto As New adodb.Recordset

'Inicio Propriedades da Classe
Public Property Get Tabela() As Integer
    Tabela = mTabela
End Property
Public Property Let Tabela(ByVal Valor As Integer)
    mTabela = Valor
End Property

Public Property Get Codigo() As String
    Codigo = mCodigo
End Property
Public Property Let Codigo(ByVal Valor As String)
    mCodigo = Valor
End Property

Public Property Get Ex() As Integer
    Ex = mEx
End Property
Public Property Let Ex(ByVal Valor As Integer)
    mEx = Ex
End Property

Public Property Get Nome() As String
    Nome = mNome
End Property
Public Property Let Nome(ByVal Valor As String)
    mNome = Valor
End Property

Public Property Get AliquotaNacional() As Currency
    AliquotaNacional = mAliquotaNacional
End Property
Public Property Let AliquotaNacional(ByVal Valor As Currency)
    mAliquotaNacional = Valor
End Property

Public Property Get AliquotaImportada() As Currency
    AliquotaImportada = mAliquotaImportada
End Property
Public Property Let AliquotaImportada(ByVal Valor As Currency)
    mAliquotaImportada = Valor
End Property
'Fim Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal pTabela As Integer, ByVal pCodigo As String) As Boolean
Dim xCondicao As String

On Error GoTo trata_erro
    
    LocalizarCodigo = False
    xCondicao = "WHERE Tabela = " & pTabela
    xCondicao = xCondicao & " AND Codigo = " & preparaTexto(pCodigo)
    gSQL = PreparaSQL(xCondicao, "ORDER BY Ex ASC")
    LocalizarCodigo = Localizar
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function

Public Function Incluir() As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO PercentualImposto (Tabela, Codigo, Ex, Nome, [Aliquota Nacional], "
    gSQL = gSQL & "[Aliquota Iternacional] ) VALUES ( "
    Call sqlNumero(1, mTabela, ", ")
    Call sqlTexto(1, mCodigo, ", ")
    Call sqlNumero(1, mEx, ", ")
    Call sqlTexto(1, mNome, ", ")
    Call sqlValor(1, mAliquotaNacional, ", ")
    Call sqlValor(1, mAliquotaImportada, " )")
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
    gSQL = "UPDATE PercentualImposto SET "
    
    Call sqlNumero(2, "Tabela = ", mTabela)
    Call sqlTexto(2, ", Codigo = ", mCodigo)
    Call sqlNumero(2, ", Ex = ", mEx)
    Call sqlTexto(2, ", Nome = ", mNome)
    Call sqlValor(2, ", [Aliquota Nacional] = ", mAliquotaNacional)
    Call sqlValor(2, ", [Aliquota Importada] = ", mAliquotaImportada)
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
    gSQL = "DELETE FROM PercentualImposto WHERE Codigo = " & nCodigo
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function

'Funções / Procedures internas
Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsPercentualImposto = New adodb.Recordset
    rsPercentualImposto.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsPercentualImposto.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsPercentualImposto.EOF Then
        AtualizaRecordset = True
    End If
End Function

Private Sub AtribuiValor()
    Tabela = rsPercentualImposto("Tabela").Value
    Codigo = rsPercentualImposto("Codigo").Value
    Ex = rsPercentualImposto("Ex").Value
    Nome = rsPercentualImposto("Nome").Value
    AliquotaNacional = rsPercentualImposto("Aliquota Nacional").Value
    AliquotaImportada = rsPercentualImposto("Aliquota Importada").Value
End Sub

Private Function Localizar() As Boolean
    Localizar = False
    If AtualizaRecordset(1) Then
        Localizar = True
        AtribuiValor
    End If
    rsPercentualImposto.Close
    Set rsPercentualImposto = Nothing
End Function

Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Tabela, Codigo, Ex, Nome, [Aliquota Nacional], [Aliquota Importada] FROM PercentualImposto" & " " & xCondicao & " " & xTipoOrdem)
End Function


