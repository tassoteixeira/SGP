VERSION 5.00
Begin VB.Form frm_libera_locacao 
   Caption         =   "Libera Manutenção"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frm_libera_locacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bd_sgp As Database
Dim tbl_dados As Table
Dim tbl_empresa As Table
Dim lNumeroHD As String
Private Sub Form_Load()
    Call ChamaDrive
    ChDir "\VB5\SGP\DATA"
    Set bd_sgp = OpenDatabase("SGP_DATA.MDB")
    Set tbl_dados = bd_sgp.OpenTable("dados")
    Set tbl_empresa = bd_sgp.OpenTable("empresas")
    Call SegurancaParaLocacao(CDate("08/09/2001"))
    With tbl_dados
        .MoveFirst
        .Edit
        ![Empresa 2] = 0
        .Update
    End With
    MsgBox "Locação Liberada", vbInformation, "Liberado!"
    End
End Sub
Function ChamaDrive() As String
    Dim dados As String
    Dim NumeroArquivo As Integer
    NumeroArquivo = FreeFile
    ChamaDrive = "C"
    On Error GoTo FileError
    Open "C:\ARQLOCAL.TXT" For Input As NumeroArquivo
    Line Input #NumeroArquivo, dados
    If Mid(dados, 1, 6) = "Drive=" Then
        ChamaDrive = Mid(dados, 7, 1)
        ChDrive ChamaDrive
    End If
    Close #NumeroArquivo
    Exit Function
FileError:
    Exit Function
End Function
Private Sub SegurancaParaLocacao(x_data As Date)
    ' Máscara da licença é:
    ' 9999      -> Ano da 1a Licença
    ' 9999      -> Número sequencial de Licença
    ' 99        -> Mes da 1a Licença
    ' 99        -> Dia da 1a Licença
    ' 999999999 -> Número de Série do HD
    lNumeroHD = DriveSerial(Left("C:", 1))
    'Cerrado Informatica  "510401793"
    'Posto Cruzeiro       "522327018"
    'Posto Pedro Ludovico "157619931"
    'Posto Colorado       "307565305"
    'Posto Goiá           "118034908"
    'Posto Solex          "-1341935335"
    'Posto Colorado       "772016339"
    'Posto Mutirão        "693377532"
    'Posto Pedro Ludovico "740826090"
    'Posto Mutirão        "642916054"
    'MsgBox "->" & lNumeroHD & "<-"
    'MsgBox Len(lNumeroHD)
    If lNumeroHD <> "510401793" And lNumeroHD <> "522327018" And lNumeroHD <> "2023215982" And lNumeroHD <> "846530529" And lNumeroHD <> "201277036" And lNumeroHD <> "1079512825" And lNumeroHD <> "157619931" And lNumeroHD <> "307565305" And lNumeroHD <> "118034908" And lNumeroHD <> "-1341935335" And lNumeroHD <> "772016339" And lNumeroHD <> "693377532" And lNumeroHD <> "740826090" And lNumeroHD <> "642916054" Then
        tbl_dados.MoveFirst
        tbl_dados.Edit
        tbl_dados![Empresa 2] = 9
        tbl_dados.Update
        tbl_dados.MoveFirst
        MsgBox "Este programa não está licenciado para esta empresa." & Chr(13) & "Sn. 2000-0037-12-01-" & lNumeroHD, vbCritical, "Atenção! Pirataria é Crime."
        End
    End If
    If Date <> x_data Then
        tbl_dados.MoveFirst
        tbl_dados.Edit
        tbl_dados![Empresa 2] = 9
        tbl_dados.Update
        tbl_dados.MoveFirst
        MsgBox "A locação deste programa está vencida!" & Chr(13) & "Efetue o pagamento e entre em contato com o suporte técnico." & Chr(13) & "Sn. 2000-0037-12-01-" & lNumeroHD, vbCritical, "Atenção! A Locação está ATRASADA."
        End
    End If
End Sub



