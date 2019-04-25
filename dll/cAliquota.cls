VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAliquota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mSerieECF As String
Private mCodigo As Integer
Private mCodigoFiscal As String
Private mNome As String
Private mAliquota As Currency

Private rsAliquota As New adodb.Recordset



'Inicio Propriedades da Classe
Public Property Get SerieECF() As String
    SerieECF = mSerieECF
End Property
Public Property Let SerieECF(ByVal Valor As String)
    mSerieECF = Valor
End Property


Public Property Get Codigo() As Integer
    Codigo = mCodigo
End Property
Public Property Let Codigo(ByVal Valor As Integer)
    mCodigo = Valor
End Property


Public Property Get CodigoFiscal() As String
    CodigoFiscal = mCodigoFiscal
End Property
Public Property Let CodigoFiscal(ByVal Valor As String)
    mCodigoFiscal = Valor
End Property


Public Property Get Nome() As String
    Nome = mNome
End Property
Public Property Let Nome(ByVal Valor As String)
    mNome = Valor
End Property


Public Property Get Aliquota() As Currency
    Aliquota = mAliquota
End Property
Public Property Let Aliquota(ByVal Valor As Currency)
    mAliquota = Valor
End Property
'Fim Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal pSerieECF As String, ByVal pCodigo As Integer) As Boolean
    Dim xCondicao As String

On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE [Serie ECF] = " & preparaTexto(pSerieECF)
    xCondicao = xCondicao & " AND Codigo = " & pCodigo
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarCodigoAliquota(ByVal pCodigo As Integer) As Boolean
    Dim xCondicao As String

On Error GoTo trata_erro

    LocalizarCodigoAliquota = False
    xCondicao = " WHERE Codigo = " & pCodigo
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigoAliquota = Localizar(1)
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarAnterior() As Boolean
Dim xSQL As String
Dim xOrdem As String
Dim xRegIgual As Boolean

On Error GoTo trata_erro
    
    LocalizarAnterior = False
    xRegIgual = False
    xSQL = "WHERE [Serie ECF] <= " & preparaTexto(mSerieECF)
    xOrdem = " ORDER BY [Serie ECF] DESC, Codigo DESC"
    gSQL = PreparaSQL(xSQL, xOrdem)
    If Localizar(20) Then
        If rsAliquota.RecordCount > 1 Then
            Do Until rsAliquota.EOF
                If xRegIgual = False Then
                    If rsAliquota("Serie ECF").Value = mSerieECF And rsAliquota("Codigo").Value = mCodigo Then
                        xRegIgual = True
                    End If
                Else
                    LocalizarAnterior = True
                    AtribuiValor
                    Exit Do
                End If
                rsAliquota.MoveNext
            Loop
        End If
    End If
    rsAliquota.Close
    Set rsAliquota = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarNome(ByVal pSerieECF As String, ByVal pNome As String) As Boolean
    Dim xCondicao As String
    
On Error GoTo trata_erro
    
    LocalizarNome = False
    xCondicao = " WHERE [Serie ECF] = " & preparaTexto(pSerieECF)
    xCondicao = xCondicao & " AND Nome = " & preparaTexto(pNome)
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarNome = Localizar(1)
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarNomeSemelhante(ByVal pNome As String) As Boolean
    Dim xCondicao As String
    
On Error GoTo trata_erro
    
    LocalizarNomeSemelhante = False
    xCondicao = " WHERE Nome LIKE " & preparaTexto("%" & pNome & "%")
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarNomeSemelhante = Localizar(1)
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
    Dim xCondicao As String

On Error GoTo trata_erro
    
    LocalizarPrimeiro = False
    gSQL = PreparaSQL("", "ORDER BY [Serie ECF] ASC, Codigo ASC")
    LocalizarPrimeiro = Localizar(1)
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarProximo() As Boolean
Dim xSQL As String
Dim xOrdem As String
Dim xRegIgual As Boolean

On Error GoTo trata_erro
    
    LocalizarProximo = False
    xRegIgual = False
    xSQL = "WHERE [Serie ECF] >= " & preparaTexto(mSerieECF)
    xOrdem = " ORDER BY [Serie ECF] ASC, Codigo ASC"
    gSQL = PreparaSQL(xSQL, xOrdem)
    If Localizar(20) Then
        If rsAliquota.RecordCount > 1 Then
            Do Until rsAliquota.EOF
                If xRegIgual = False Then
                    If rsAliquota("Serie ECF").Value = mSerieECF And rsAliquota("Codigo").Value = mCodigo Then
                        xRegIgual = True
                    End If
                Else
                    LocalizarProximo = True
                    AtribuiValor
                    Exit Do
                End If
                rsAliquota.MoveNext
            Loop
        End If
    End If
    rsAliquota.Close
    Set rsAliquota = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo() As Boolean
On Error GoTo trata_erro
    
    LocalizarUltimo = False
    gSQL = PreparaSQL("", "ORDER BY [Serie ECF] DESC, Codigo DESC")
    LocalizarUltimo = Localizar(1)
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean
    Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Aliquota ( Codigo, [Codigo Fiscal], Nome, [Aliquota do Imposto], [Serie ECF] ) VALUES ( "
    Call sqlNumero(1, mCodigo, ", ")
    Call sqlTexto(1, mCodigoFiscal, ", ")
    Call sqlTexto(1, mNome, ", ")
    Call sqlValor(1, mAliquota, ", ")
    Call sqlTexto(1, mSerieECF, " )")
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    Incluir = True
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal pSerieECF As String, ByVal pCodigo As Integer) As Boolean

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Aliquota SET "
    Call sqlTexto(2, "Nome = ", mNome)
    Call sqlTexto(2, ", [Codigo Fiscal] = ", mCodigoFiscal)
    Call sqlValor(2, ", [Aliquota do Imposto] = ", mAliquota)
    Call sqlTexto(2, ", [Serie ECF] = ", mSerieECF)
    gSQL = gSQL & " WHERE [Serie ECF] = " & preparaTexto(pSerieECF)
    gSQL = gSQL & " AND Codigo = " & pCodigo
    gConn.Execute gSQL, Alterar, adCmdText + adExecuteNoRecords
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal pSerieECF As String, ByVal pCodigo As Integer) As Boolean

On Error GoTo Err_Excluir
            
    Excluir = False
    gSQL = "DELETE FROM Aliquota"
    gSQL = gSQL & " WHERE [Serie ECF] = " & preparaTexto(pSerieECF)
    gSQL = gSQL & " AND Codigo = " & pCodigo
    gConn.Execute gSQL, Excluir, adCmdText + adExecuteNoRecords
    Exit Function

Err_Excluir:
End Function



'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Codigo, [Codigo Fiscal], Nome, [Aliquota do Imposto], [Serie ECF] FROM Aliquota" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsAliquota = New adodb.Recordset
    rsAliquota.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsAliquota.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsAliquota.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Codigo = rsAliquota("Codigo").Value
    CodigoFiscal = rsAliquota("Codigo Fiscal").Value
    Nome = rsAliquota("Nome").Value
    Aliquota = rsAliquota("Aliquota do Imposto").Value
    SerieECF = rsAliquota("Serie ECF").Value
End Sub


Private Function Localizar(ByVal pQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(pQtdRegistro) Then
        Localizar = True
        AtribuiValor
    End If
    rsAliquota.Close
    Set rsAliquota = Nothing
End Function
