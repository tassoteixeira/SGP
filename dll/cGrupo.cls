VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCodigo As Integer
Private mNome As String
Private mCstEntrada As String
Private mCfopEntrada As String
Private mCstSaida As String
Private mCfopSaida As String
Private mCstCofinsEntrada As String
Private mAliquotaCofins As Currency
Private mCstPisEntrada As String
Private mAliquotaPis As Currency
Private mCstCofinsSaida As String
Private mCstPisSaida As String


Private rsGrupo As New adodb.Recordset
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


Public Property Get CstEntrada() As String
    CstEntrada = mCstEntrada
End Property
Public Property Let CstEntrada(ByVal Valor As String)
    mCstEntrada = Valor
End Property


Public Property Get CfopEntrada() As String
    CfopEntrada = mCfopEntrada
End Property
Public Property Let CfopEntrada(ByVal Valor As String)
    mCfopEntrada = Valor
End Property


Public Property Get CstSaida() As String
    CstSaida = mCstSaida
End Property
Public Property Let CstSaida(ByVal Valor As String)
    mCstSaida = Valor
End Property


Public Property Get CfopSaida() As String
    CfopSaida = mCfopSaida
End Property
Public Property Let CfopSaida(ByVal Valor As String)
    mCfopSaida = Valor
End Property


Public Property Get CstCofinsEntrada() As String
    CstCofinsEntrada = mCstCofinsEntrada
End Property
Public Property Let CstCofinsEntrada(ByVal Valor As String)
    mCstCofinsEntrada = Valor
End Property


Public Property Get AliquotaCofins() As Currency
    AliquotaCofins = mAliquotaCofins
End Property
Public Property Let AliquotaCofins(ByVal Valor As Currency)
    mAliquotaCofins = Valor
End Property


Public Property Get CstPisEntrada() As String
    CstPisEntrada = mCstPisEntrada
End Property
Public Property Let CstPisEntrada(ByVal Valor As String)
    mCstPisEntrada = Valor
End Property


Public Property Get AliquotaPis() As Currency
    AliquotaPis = mAliquotaPis
End Property
Public Property Let AliquotaPis(ByVal Valor As Currency)
    mAliquotaPis = Valor
End Property


Public Property Get CstCofinsSaida() As String
    CstCofinsSaida = mCstCofinsSaida
End Property
Public Property Let CstCofinsSaida(ByVal Valor As String)
    mCstCofinsSaida = Valor
End Property


Public Property Get CstPisSaida() As String
    CstPisSaida = mCstPisSaida
End Property
Public Property Let CstPisSaida(ByVal Valor As String)
    mCstPisSaida = Valor
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
Dim lRecordsAffected As Long

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Grupo ( Codigo, Nome, [CST de Entrada], [CFOP de Entrada], [CST de Saida], "
    gSQL = gSQL & "[CFOP de Saida], [CST COFINS Entrada], [Aliquota COFINS], [CST PIS Entrada], [Aliquota PIS], "
    gSQL = gSQL & "[CST COFINS Saida], [CST PIS Saida] ) VALUES ( "
    Call sqlNumero(1, mCodigo, ", ")
    Call sqlTexto(1, mNome, ", ")
    Call sqlTexto(1, mCstEntrada, ", ")
    Call sqlTexto(1, mCfopEntrada, ", ")
    Call sqlTexto(1, mCstSaida, ", ")
    Call sqlTexto(1, mCfopSaida, ", ")
    Call sqlTexto(1, mCstCofinsEntrada, ", ")
    Call sqlValor(1, mAliquotaCofins, ", ")
    Call sqlTexto(1, mCstPisEntrada, ", ")
    Call sqlValor(1, mAliquotaPis, ", ")
    Call sqlTexto(1, mCstCofinsSaida, ", ")
    Call sqlTexto(1, mCstPisSaida, " )")
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
    Call sqlTexto(2, "Nome = ", mNome)
    Call sqlTexto(2, ", [CST de Entrada] = ", mCstEntrada)
    Call sqlTexto(2, ", [CFOP de Entrada] = ", mCfopEntrada)
    Call sqlTexto(2, ", [CST de Saida] = ", mCstSaida)
    Call sqlTexto(2, ", [CFOP de Saida] = ", mCfopSaida)
    Call sqlTexto(2, ", [CST COFINS Entrada] = ", mCstCofinsEntrada)
    Call sqlValor(2, ", [Aliquota COFINS] = ", mAliquotaCofins)
    Call sqlTexto(2, ", [CST PIS Entrada] = ", mCstPisEntrada)
    Call sqlValor(2, ", [Aliquota PIS] = ", mAliquotaPis)
    Call sqlTexto(2, ", [CST COFINS Saida] = ", mCstCofinsSaida)
    Call sqlTexto(2, ", [CST PIS Saida] = ", mCstPisSaida)
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
    gSQL = "DELETE FROM Grupo WHERE Codigo = " & nCodigo
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Excluir = True
    End If
    Exit Function

Err_Excluir:
End Function



'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Codigo, Nome, [CST de Entrada], [CFOP de Entrada], [CST de Saida], [CFOP de Saida], [CST COFINS Entrada], [Aliquota COFINS], [CST PIS Entrada], [Aliquota PIS], [CST COFINS Saida], [CST PIS Saida] FROM Grupo" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsGrupo = New adodb.Recordset
    rsGrupo.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsGrupo.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsGrupo.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Codigo = rsGrupo("Codigo").Value
    Nome = rsGrupo("Nome").Value
    CstEntrada = rsGrupo("CST de Entrada").Value
    CfopEntrada = rsGrupo("CFOP de Entrada").Value
    CstSaida = rsGrupo("CST de Saida").Value
    CfopSaida = rsGrupo("CFOP de Saida").Value
    CstCofinsEntrada = rsGrupo("CST COFINS Entrada").Value
    AliquotaCofins = rsGrupo("Aliquota COFINS").Value
    CstPisEntrada = rsGrupo("CST PIS Entrada").Value
    AliquotaPis = rsGrupo("Aliquota PIS").Value
    CstCofinsSaida = rsGrupo("CST COFINS Saida").Value
    CstPisSaida = rsGrupo("CST PIS Saida").Value
End Sub


Private Function Localizar() As Boolean
    Localizar = False
    If AtualizaRecordset(1) Then
        Localizar = True
        AtribuiValor
    End If
    rsGrupo.Close
    Set rsGrupo = Nothing
End Function
