Attribute VB_Name = "ecf_mecaf"
Declare Function OpenCif Lib "ECF32M.DLL" () As Long
Declare Sub CloseCif Lib "ECF32M.DLL" ()
Declare Function AbreCupomFiscal Lib "ECF32M.DLL" () As Long
Declare Function LeituraX Lib "ECF32M.DLL" (ByVal RelGer As Byte) As Long
Declare Function ReducaoZ Lib "ECF32M.DLL" (ByVal RelGer As Byte) As Long
Declare Function ProgramaHorarioVerao Lib "ECF32M.DLL" (ByVal hv As Byte) As Long
Declare Function ProgramaLegenda Lib "ECF32M.DLL" (ByVal reg As String, ByVal leg As String) As Long
Declare Function TransStatus Lib "ECF32M.DLL" (ByVal BitTest As Long, ByVal BufStat As String) As Long
Declare Function TransTotCont Lib "ECF32M.DLL" () As Long
Declare Function TransDataHora Lib "ECF32M.DLL" () As Long
Declare Function TotalizarCupom Lib "ECF32M.DLL" (ByVal oper As Byte, ByVal toper As Byte, ByVal valor As String, ByVal legendaOp As String) As Long
Declare Function Pagamento Lib "ECF32M.DLL" (ByVal reg As String, ByVal vpgto As String, ByVal subtr As Byte) As Long
Declare Function LeMemFiscalData Lib "ECF32M.DLL" (ByVal datai As String, ByVal dataf As String, ByVal res As Byte) As Long
Declare Function ObtemRetorno Lib "ECF32M.DLL" (ByVal buf_ret As String) As Long
Declare Function VendaItem Lib "ECF32M.DLL" (ByVal fmt As Byte, ByVal qtd As String, ByVal punit As String, ByVal trib As String, ByVal TDesc As Byte, ByVal valor As String, ByVal unid As String, ByVal cod As String, ByVal ex As Byte, ByVal descr As String, ByVal legendaOp As String) As Long
Declare Function CancelamentoItem Lib "ECF32M.DLL" (ByVal numitem As String) As Long
Declare Function DescontoItem Lib "ECF32M.DLL" (ByVal toper As Byte, ByVal valor As String, ByVal legop As String) As Long
Declare Function FechaCupomFiscal Lib "ECF32M.DLL" (ByVal tam_msg As String, ByVal msg As String) As Long
Declare Function CancelaCupomFiscal Lib "ECF32M.DLL" () As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Function TrataRetorno(lngRet As Long) As String
    Dim strMsg As String
    Dim lngBufRet As Long
    Dim strBufRet As String
    
    strBufRet = ObtemRet(lngRet)
    
    strMsg = TraduzCodigoRetorno(lngRet)
'    MsgBox " " + Str(lngRet) + " - " + strMsg
    
    TrataRetorno = strBufRet
    
End Function
Public Function ObtemRet(lngRet As Long) As String
    Const MaxBuf = 2000
    Dim strBufRet As String * MaxBuf
    Dim intCont As Integer

    If lngRet = CIF_OK Then
        For intCont = 1 To 40
            lngRet = ObtemRetorno(strBufRet)
            If (lngRet <> -97) Then
                Exit For
            End If
            Sleep 1000
        Next intCont
    End If
    ObtemRet = strBufRet
End Function
Function TraduzCodigoRetorno(ByVal intretorno As Integer) As String
    Dim strMsg As String
    Select Case intretorno
    
        '---------------------------
        ' Codigo de retorno dos comandos da impressora
        '
        Case -1
            strMsg = "Cabeçalho contém caracteres inválidos"
        Case -2
            strMsg = "Comando inexistente"
        Case -3
            strMsg = "Valor não numérico em campo numérico"
        Case -4
            strMsg = "Valor fora da faixa entre 20h e 7Fh"
        Case -5
            strMsg = "Campo deve iniciar com @, & ou %"
        Case -6
            strMsg = "Campo deve iniciar com $, # ou ?"
        Case -7
            strMsg = "O intervalo é inconsistente. No caso de datas, valores anteriores a " & _
                   "01/01/95 serão consideradas como ano 2000 a 2094"
        Case -9
            strMsg = "A string TOTAL não é aceita"
        Case -10
            strMsg = "A sintaxe do comando está errada"
        Case -11
            strMsg = "Excedeu número máximo de linhas permitidas pelo comando"
        Case -12
            strMsg = "O terminador enviado não está obedecendo o protocolo de comunicação"
        Case -13
            strMsg = "O checksum está incorreto"
        Case -15
            strMsg = "A situação tributária deve iniciar com T, F, I ou N"
        Case -16
            strMsg = "Data inválida"
        Case -17
            strMsg = "Hora inválida"
        Case -18
            strMsg = "Alíquota não programada ou fora do intervalo"
        Case -19
            strMsg = "O campo de sinal está incorreto"
        Case -20
            strMsg = "Comando só aceito em Intervenção Fiscal"
        Case -22
            strMsg = "É necessário abrir o Cupom Fiscal"
        Case -23
            strMsg = "Comando não aceito durante Cupom Fiscal"
        Case -24
            strMsg = "É necessário abrir Cupom Não Fiscal"
        Case -25
            strMsg = "Comando não aceito durante Cupom Não Fiscal"
        Case -26
            strMsg = "O relógio já está em horário de verão"
        Case -27
            strMsg = "O relógio não está em horário de verão"
        Case -28
            strMsg = "Necessário realizar Redução Z"
        Case -29
            strMsg = "Fechamento do dia (Redução Z) já executado"
        Case -30
            strMsg = "Necessário programar legenda"
        Case -31
            strMsg = "Item inexistente ou já cancelado"
        Case -32
            strMsg = "O cupom anterior não pode ser cancelado"
        Case -33
            strMsg = "Detectado falta de papel. Verifique a impressora."
        Case -36
            strMsg = "Necessário programar os dados do estabelecimento"
        Case -37
            strMsg = "Necessário realizar Intervenção Fiscal."
        Case -38
        strMsg = "Memória Fiscal não permite mais realizar vendas. Apenas é possível realizar LeituraX " & _
               "ou Leitura da Memória Fiscal."
        Case -39
            strMsg = "Memória Fiscal não permite mais realizar vendas. Apenas é possível realizar LeituraX " & _
                   "ou Leitura da Memória Fiscal, deve haver algum problema na memória NOVRAM. Será " & _
                   "necessário realizar Intervenção Fiscal."
        Case -40
            strMsg = "Necessário programar a data do relógio"
        Case -41
            strMsg = "Número máximo de itens por cupom ultrapassado"
        Case -42
            strMsg = "Já foi realizado o Ajuste de Hora Diário"
        Case -43
            strMsg = "Comando válido ainda em execução"
        Case -44
            strMsg = "Está em estado de Impressão de Cheques"
        Case -45
            strMsg = "Não está em estado de Impressão de Cheques"
        Case -46
            strMsg = "Necessário inserir o cheque"
        Case -47
            strMsg = "Necessário inserir nova bobina"
        Case -48
            strMsg = "Necessário executar uma Leitura X"
        Case -49
            strMsg = "Detectado algum problema na impressora (Paper jam, sobretensão, etc)."
        Case -50
            strMsg = "Cupom já totalizado"
        Case -51
            strMsg = "Necessário totalizar cupom antes de fechar"
        Case -52
            strMsg = "Necessário finalizar Cupom com comando correto"
        Case -53
            strMsg = "Ocorreu erro de gravação na Memória Fiscal"
        Case -54
            strMsg = "Excedeu número máximo de estabelecimentos"
        Case -55
            strMsg = "Memória fiscal não inicializada"
        Case -56
            strMsg = "Ultrapassou valor do pagamento"
        Case -57
            strMsg = "Registrador não programado ou troco já realizado"
        Case -58
            strMsg = "Falta completar valor do pagamento"
        Case -59
            strMsg = "Campo somente de caracteres não numéricos"
        Case -60
            strMsg = "Excedeu campo máximo de caracteres"
        Case -61
            strMsg = "Troco não realizado"
        Case -62
            strMsg = "Comando desabilitado"
            
        '---------------------------
        ' Codigo de retorno de funcoes da DLL
        '
        Case CIF_OK
            strMsg = "Operação efetuada com sucesso"
        Case CIF_PPAPEL
            strMsg = "Sucesso, detectado pouco papel"
        Case CIF_CANCCUP
            strMsg = "Sucesso, cancelando cupom"
        Case CIF_CUPNF
            strMsg = "Sucesso, abrindo cupom rel gerencial"
        Case CIF_ERR
            strMsg = "Falha geral na execução da DLL"
        Case CIF_EMEXECUCAO
            strMsg = "Comando válido ainda em execução"
        Case CIF_ERR_CONFIG
            strMsg = "Erro no arquivo CIF.INI"
        Case CIF_ERR_SERIAL
            strMsg = "Erro na abertura da serial"
        Case CIF_ERR_SYS
            strMsg = "Falha na alocação de recursos do Windows"
        Case CIF_ERR_ANSWER
            strMsg = "Retorno nao reconhecido"
        Case CIF_ERR_READSER
            strMsg = "Falha na leitura da serial"
        Case CIF_ERR_TEMP
            strMsg = "Temperatura da cabeça de impressão alta"
        Case CIF_ERR_PPAPEL
            strMsg = "Pouco papel"
        Case CIF_IRRECUPERAVEL
            strMsg = "Erro irrecuperável"
        Case CIF_ERR_MECANICO
            strMsg = "Erro mecânico"
        Case CIF_ERR_TABERTA
            strMsg = "Tampa aberta"
        Case CIF_SEMRETORNO
            strMsg = "Operação sem retorno"
        Case CIF_OVERFLOW
            strMsg = "Buffer overflow. Tamanho da mensagem enviada pelo ECF é maior do que o buffer fornecido pela aplicação"
        Case CIF_TIMEOUT
            strMsg = "TimeOut na execucao do comando"
        Case Else
            strMsg = "Código de retorno inexistente"
    End Select
    TraduzCodigoRetorno = strMsg
'    MsgBox strMsg
End Function
Function RetornaBStatus(ByVal NumBit As Long) As Boolean
    Dim lngRet As Long
    Dim strBuffer As String * 40
    Dim strBufferFormatado As String * 40
    Dim strByte1  As String * 8
    Dim strByte2  As String * 8
    Dim strByte3  As String * 8
    Dim strByte4  As String * 8
    Dim strByte5  As String * 8
    RetornaBStatus = False
    strBuffer = String(MaxSize, 0)
    lngRet = TransStatus(0, strBuffer)
    If lngRet <> CIF_OK Then
        Exit Function
    End If
    strByte1 = StrReverse(Mid(strBuffer, 1, 8))
    strByte2 = StrReverse(Mid(strBuffer, 9, 8))
    strByte3 = StrReverse(Mid(strBuffer, 17, 8))
    strByte4 = StrReverse(Mid(strBuffer, 25, 8))
    strByte5 = StrReverse(Mid(strBuffer, 33, 8))
    strBufferFormatado = strByte1 & strByte2 & strByte3 & strByte4 & strByte5
    If Mid$(strBufferFormatado, NumBit, 1) = "1" Then
        RetornaBStatus = True
    End If
End Function

