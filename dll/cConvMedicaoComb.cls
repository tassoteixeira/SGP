VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cConvMedicaoComb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mEmpresa As Integer
Private mMedida As Integer
Private mMedicaoTanque10 As Currency
Private mMedicaoTanque15 As Currency
Private mMedicaoTanque20 As Currency
Private mMedicaoTanque30 As Currency

Private rsConversaoMedicaoCombustivel As New adodb.Recordset


'Inicio das Propriedades da Classe
Public Property Get Empresa() As Integer
    Empresa = mEmpresa
End Property
Public Property Let Empresa(ByVal Valor As Integer)
    mEmpresa = Valor
End Property


Public Property Get Medida() As Integer
    Medida = mMedida
End Property
Public Property Let Medida(ByVal Valor As Integer)
    mMedida = Valor
End Property


Public Property Get MedicaoTanque10() As Currency
    MedicaoTanque10 = mMedicaoTanque10
End Property
Public Property Let MedicaoTanque10(ByVal Valor As Currency)
    mMedicaoTanque10 = Valor
End Property


Public Property Get MedicaoTanque15() As Currency
    MedicaoTanque15 = mMedicaoTanque15
End Property
Public Property Let MedicaoTanque15(ByVal Valor As Currency)
    mMedicaoTanque15 = Valor
End Property


Public Property Get MedicaoTanque20() As Currency
    MedicaoTanque20 = mMedicaoTanque20
End Property
Public Property Let MedicaoTanque20(ByVal Valor As Currency)
    mMedicaoTanque20 = Valor
End Property


Public Property Get MedicaoTanque30() As Currency
    MedicaoTanque30 = mMedicaoTanque30
End Property
Public Property Let MedicaoTanque30(ByVal Valor As Currency)
    mMedicaoTanque30 = Valor
End Property
'Fim das Propriedades da Classe


'Inicio Métodos da Classe
Public Function LocalizarCodigo(ByVal nEmpresa As Integer, nMedida As Integer) As Boolean
Dim xCondicao As String

On Error GoTo trata_erro

    LocalizarCodigo = False
    xCondicao = " WHERE Empresa = " & nEmpresa
    xCondicao = xCondicao & " AND Medida = " & nMedida
    gSQL = PreparaSQL(xCondicao, "")
    LocalizarCodigo = Localizar(1)
    If LocalizarCodigo Then
        AtribuiValor
    End If
    rsConversaoMedicaoCombustivel.Close
    Set rsConversaoMedicaoCombustivel = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarAnterior() As Boolean
Dim xCondicao As String

On Error GoTo trata_erro

    LocalizarAnterior = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND Medida < " & mMedida
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa DESC, Medida DESC")
    LocalizarAnterior = Localizar(1)
    If LocalizarAnterior Then
        AtribuiValor
    End If
    rsConversaoMedicaoCombustivel.Close
    Set rsConversaoMedicaoCombustivel = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarPrimeiro() As Boolean
Dim xCondicao As String

On Error GoTo trata_erro

    LocalizarPrimeiro = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND Medida > 0"
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa ASC, Medida ASC")
    LocalizarPrimeiro = Localizar(1)
    If LocalizarPrimeiro Then
        AtribuiValor
    End If
    rsConversaoMedicaoCombustivel.Close
    Set rsConversaoMedicaoCombustivel = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarProximo() As Boolean
Dim xCondicao As String

On Error GoTo trata_erro

    LocalizarProximo = False
    xCondicao = " WHERE Empresa = " & mEmpresa
    xCondicao = xCondicao & " AND Medida > " & mMedida
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa ASC, Medida ASC")
    LocalizarProximo = Localizar(1)
    If LocalizarProximo Then
        AtribuiValor
    End If
    rsConversaoMedicaoCombustivel.Close
    Set rsConversaoMedicaoCombustivel = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function LocalizarUltimo(ByVal nEmpresa As Integer) As Boolean
Dim xCondicao As String

On Error GoTo trata_erro

    LocalizarUltimo = False
    xCondicao = " WHERE Empresa = " & nEmpresa
    xCondicao = xCondicao & " AND Medida < 9999"
    gSQL = PreparaSQL(xCondicao, "ORDER BY Empresa DESC, Medida DESC")
    LocalizarUltimo = Localizar(1)
    If LocalizarUltimo Then
        AtribuiValor
    End If
    rsConversaoMedicaoCombustivel.Close
    Set rsConversaoMedicaoCombustivel = Nothing
    Exit Function

trata_erro:
    MsgBox Err.Number & " - " & Err.Description
End Function


Public Function Incluir() As Boolean

On Error GoTo Err_Incluir

    Incluir = False
    gSQL = "INSERT INTO Conversao_Medicao_Combustivel ( Empresa, Medida, [Medicao Tanque 10], [Medicao Tanque 15], [Medicao Tanque 20], [Medicao Tanque 30] ) VALUES ( "
    Call sqlNumero(1, mEmpresa, ", ")
    Call sqlNumero(1, mMedida, ", ")
    Call sqlValor(1, mMedicaoTanque10, ", ")
    Call sqlValor(1, mMedicaoTanque15, ", ")
    Call sqlValor(1, mMedicaoTanque20, ", ")
    Call sqlValor(1, mMedicaoTanque30, " )")
    gConn.Execute gSQL, Incluir, adCmdText + adExecuteNoRecords
    Incluir = True
    Exit Function

Err_Incluir:
End Function


Public Function Alterar(ByVal nEmpresa As Integer, nMedida As Integer) As Boolean
Dim lRecordsAffected As Long

On Error GoTo Err_Alterar

    Alterar = False
    gSQL = "UPDATE Conversao_Medicao_Combustivel SET "
    Call sqlNumero(2, "Empresa = ", mEmpresa)
    Call sqlNumero(2, ", Medida = ", mMedida)
    Call sqlValor(2, ", [Medicao Tanque 10] = ", mMedicaoTanque10)
    Call sqlValor(2, ", [Medicao Tanque 15] = ", mMedicaoTanque15)
    Call sqlValor(2, ", [Medicao Tanque 20] = ", mMedicaoTanque20)
    Call sqlValor(2, ", [Medicao Tanque 30] = ", mMedicaoTanque30)
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND Medida = " & nMedida
    gConn.Execute gSQL, lRecordsAffected, adCmdText + adExecuteNoRecords
    If lRecordsAffected > 0 Then
        Alterar = True
    End If
    Exit Function

Err_Alterar:
End Function


Public Function Excluir(ByVal nEmpresa As Integer, nMedida As Integer) As Boolean

On Error GoTo Err_Excluir

    Excluir = False
    gSQL = "DELETE FROM Conversao_Medicao_Combustivel"
    gSQL = gSQL & " WHERE Empresa = " & nEmpresa
    gSQL = gSQL & " AND Medida = " & nMedida
    gConn.Execute gSQL, Excluir, adCmdText + adExecuteNoRecords
    Excluir = True
    Exit Function

Err_Excluir:
End Function


'Funções / Procedures internas
Private Function PreparaSQL(ByVal xCondicao As String, ByVal xTipoOrdem As String) As String
    PreparaSQL = Trim("SELECT Empresa, Medida, [Medicao Tanque 10], [Medicao Tanque 15], [Medicao Tanque 20], [Medicao Tanque 30] FROM Conversao_Medicao_Combustivel" & " " & xCondicao & " " & xTipoOrdem)
End Function


Private Function AtualizaRecordset(ByVal xQtdRegistro As Integer) As Boolean
    Dim i As Integer
    AtualizaRecordset = False
    Set rsConversaoMedicaoCombustivel = New adodb.Recordset
    rsConversaoMedicaoCombustivel.CursorLocation = adUseClient
    i = Len(gSQL)
    If xQtdRegistro > 0 Then
        gSQL = Mid(gSQL, 1, 6) & " TOP " & xQtdRegistro & Mid(gSQL, 7, i - 6)
    End If
    rsConversaoMedicaoCombustivel.Open gSQL, gConn, adOpenForwardOnly, adLockReadOnly
    If Not rsConversaoMedicaoCombustivel.EOF Then
        AtualizaRecordset = True
    End If
End Function


Private Sub AtribuiValor()
    Empresa = rsConversaoMedicaoCombustivel("Empresa").Value
    Medida = rsConversaoMedicaoCombustivel("Medida").Value
    MedicaoTanque10 = rsConversaoMedicaoCombustivel("Medicao Tanque 10").Value
    MedicaoTanque15 = rsConversaoMedicaoCombustivel("Medicao Tanque 15").Value
    MedicaoTanque20 = rsConversaoMedicaoCombustivel("Medicao Tanque 20").Value
    MedicaoTanque30 = rsConversaoMedicaoCombustivel("Medicao Tanque 30").Value
End Sub


Private Function Localizar(ByVal xQtdRegistro As Integer) As Boolean
    Localizar = False
    If AtualizaRecordset(xQtdRegistro) Then
        Localizar = True
    End If
End Function

