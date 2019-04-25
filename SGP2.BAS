Attribute VB_Name = "sgp_variaveis_globais2"
Option Explicit
Global bd_sgp As Database
Global bd_sgp_b As Database
Global bd_sgp_m As Database
Global bd_disco As Database

Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const CB_SHOWDROPDOWN = &H14F

Global DefaultWorkspace As Workspace
Global tabela As TableDef
Global campo As Field
Global indice As Index

Global tbl_acm_venda_combustivel As Table
Global tbl_baixa_duplicata_receber As Table
Global tbl_baixa_nota As Table
Global tbl_baixa_pagar As Table
Global tbl_banco As Table
Global tbl_bomba As Table
Global tbl_cheque As Table
Global tbl_cheque_baixado As Table
Global tbl_cheque_emitido As Table
Global tbl_cheque_folha As Table
Global tbl_cliente As Table
Global tbl_cliente_conveniado As Table
Global tbl_combustivel As Table
Global tbl_configuracao As Table
Global tbl_conta As Table
Global tbl_conta_bancaria As Table
Global tbl_contas_pagar As Table
Global tbl_convenio As Table
Global tbl_dados As Table
Global tbl_duplicata_receber As Table
Global tbl_empresa As Table
Global tbl_entrada_comb As Table
Global tbl_estoque As Table
Global tbl_fornecedor As Table
Global tbl_funcionario As Table
Global tbl_grupo As Table
Global tbl_historico As Table
Global tbl_local_cobranca As Table
Global tbl_med_comb As Table
Global tbl_menu As Table
Global tbl_mov_bomba As Table
Global tbl_movimento_nota As Table
Global tbl_mov_historico As Table
Global tbl_pagina_lmc As Table
Global tbl_pedido_comb As Table
Global tbl_produto As Table
Global tbl_programa As Table
Global tbl_saldo_conta As Table
Global tbl_setup As Table
Global tbl_tipodoc As Table
Global tbl_usuario As Table
Global tbl_vencimento As Table

Global g_nivel_acesso As Integer
Global g_lmc As Integer
Global g_usuario As Integer
Global g_empresa As Integer
Global g_nome_empresa As String
Global g_caixa_unificado As Boolean
Global g_cidade_empresa As String
Global g_nome_conta As String
Global g_historico As Long
Global g_banco As Integer
Global g_tipodoc As Integer
Global g_tamanho_impressora As Integer
Global g_conta_bancaria As String * 10
Global g_data As Date
Global g_data_def As Date
Global g_valor As Currency
Public g_nome_tabela As String
Public g_nome_usuario
Public g_string As String

Public g_cfg_data_i As Date
Public g_cfg_data_f As Date
Public g_cfg_empresa_i As Integer
Public g_cfg_empresa_f As Integer
Public g_cfg_periodo_i As String * 1
Public g_cfg_periodo_f As String * 1

Global Const ASC_AA = 65, ASC_ZZ = 90
Global Const ASC_a = 97, ASC_z = 122
Global Const ASC_0 = 48, ASC_9 = 57
Function SelecionaImpressoraEpson(x_form As Form) As Boolean
    g_tamanho_impressora = 132
    SelecionaImpressoraEpson = True
    If UCase(Printer.DeviceName) Like "*300*" Or UCase(Printer.DeviceName) Like "*810*" Then
        g_tamanho_impressora = 80
    End If
    If UCase(Printer.DeviceName) Like "*EPSON*" Then
        Exit Function
    End If
    x_form.CommonDialog1.ShowPrinter
    x_form.CommonDialog1.PrinterDefault = True
    If UCase(Printer.DeviceName) Like "*300*" Or UCase(Printer.DeviceName) Like "*810*" Then
        g_tamanho_impressora = 80
    End If
    If UCase(Printer.DeviceName) <> "*EPSON*" Then
        Exit Function
    End If
    MsgBox "Não será possível imprimir na impressora " & Printer.DeviceName & ".", 48
    SelecionaImpressoraEpson = False
End Function
Function SelecionaImpressoraHP(x_form As Form) As Boolean
    SelecionaImpressoraHP = True
    If UCase(Printer.DeviceName) Like "*HP*" Then
        Exit Function
    End If
    x_form.CommonDialog1.ShowPrinter
    x_form.CommonDialog1.PrinterDefault = True
    If UCase(Printer.DeviceName) Like "*HP*" Then
        Exit Function
    End If
    MsgBox "Não será possível imprimir na impressora " & Printer.DeviceName & ".", 48
    SelecionaImpressoraHP = False
End Function
Sub ValidaInteiro(tecla As Integer)
    Dim char As String
    char = Chr$(tecla)
    If char < "0" Or char > "9" Then
        If tecla = 8 Then
            Exit Sub
        End If
        tecla = 0
    End If
End Sub
Sub ValidaLetraMaiuscula(tecla As Integer)
    Dim char As String
    char = UCase(Chr$(tecla))
    tecla = Asc(char)
    If char < "A" Or char > "Z" Then
        If tecla = 8 Then
            Exit Sub
        End If
        tecla = 0
    End If
End Sub
Function BioCriaImprime() As String
    BioCriaImprime = Format(Day(Date), "00") & Format(Time, "hhmmss") & ".TMP"
    Open BioCriaImprime For Output As #1
End Function
Sub BioImprime(x_string As String)
    Print #1, x_string
End Sub
Sub BioFechaImprime()
    Close #1
End Sub
Function BuscaArquivoConfiguracao(x_numero As Integer) As String
    Dim x_string As String
    Dim i As Integer
    Dim i2 As Integer
    Dim NumeroArquivo As Integer
    NumeroArquivo = FreeFile
    On Error GoTo FileError
    i2 = 0
    BuscaArquivoConfiguracao = "BRADESCO"
    Open "\VB5\SGP\DATA\configuracao.cfg" For Input As NumeroArquivo
    For i = 0 To 100
        Input #NumeroArquivo, x_string
        If Mid(x_string, 1, 1) <> "@" Then
            i2 = i2 + 1
        End If
        If i2 = x_numero Then
            BuscaArquivoConfiguracao = Trim(x_string)
            Exit For
        End If
    Next
    Close NumeroArquivo
    Exit Function
FileError:
    If Err = 62 Then
        Close NumeroArquivo
        Exit Function
    End If
    MsgBox "O arquivo de configuração CONFIGURACAO.CFG não foi encontrado!", vbCritical, "Erro de Configuração!"
    Exit Function
End Function
Function CalculaDigito5(x_codigo As Long) As Integer
    Dim x_numero As String
    Dim x_digito As String
    Dim soma As Integer
    Dim tabela As Integer
    x_numero = Format(x_codigo, "00000")
    soma = Mid(x_numero, 1, 1) * 6 + Mid(x_numero, 2, 1) * 5 + Mid(x_numero, 3, 1) * 4 + Mid(x_numero, 4, 1) * 3 + Mid(x_numero, 5, 1) * 2
    tabela = 0
    Do Until soma < tabela
        tabela = tabela + 11
    Loop
    x_digito = tabela - soma
    If x_digito = 0 Then
        x_digito = 1
    End If
    x_digito = Format(x_digito, "00000")
    x_digito = Mid(x_digito, 5, 1)
    CalculaDigito5 = Val(x_digito)
End Function
Function CalculaDigitoCGC(x_cgc As String) As Boolean
    Dim x_soma As Currency
    Dim i As Integer
    Dim x_digito_1 As Integer
    Dim x_digito_2 As Integer
    x_soma = 0
    For i = 1 To 4
        x_soma = x_soma + (Val(Mid(x_cgc, i, 1)) * (6 - i))
    Next
    For i = 5 To 12
        x_soma = x_soma + (Val(Mid(x_cgc, i, 1)) * (14 - i))
    Next
    x_soma = (x_soma - Val(x_soma / 11) * 11)
    If x_soma = 1 Or x_soma = 0 Then
        x_digito_1 = 0
    Else
        x_digito_1 = 11 - x_soma
    End If
    x_soma = 0
    For i = 1 To 5
        x_soma = x_soma + (Val(Mid(x_cgc, i, 1)) * (7 - i))
    Next
    For i = 6 To 12
        x_soma = x_soma + (Val(Mid(x_cgc, i, 1)) * (15 - i))
    Next
    x_soma = x_soma + (x_digito_1 * 2)
    x_soma = (x_soma - Val(x_soma / 11) * 11)
    If x_soma = 1 Or x_soma = 0 Then
        x_digito_2 = 0
    Else
        x_digito_2 = 11 - x_soma
    End If
    If Mid(x_cgc, 13, 1) = x_digito_1 And Mid(x_cgc, 14, 1) = x_digito_2 Then
        CalculaDigitoCGC = True
    Else
        CalculaDigitoCGC = False
        MsgBox "O CGC: " & Mid(x_cgc, 1, 2) & "." & Mid(x_cgc, 3, 3) & "." & Mid(x_cgc, 6, 3) & "/" & Mid(x_cgc, 9, 4) & "-" & Mid(x_cgc, 13, 2) & " está com o dígito errado!" & Chr(10) & "O dígito correto é: " & Chr(34) & x_digito_1 & x_digito_2 & Chr(34), vbInformation, "Dígito de CGC Inválido!"
    End If
End Function
Function CalculaDigitoCPF(x_cpf As String) As Boolean
    Dim x_soma As Currency
    Dim i As Integer
    Dim x_digito_1 As Integer
    Dim x_digito_2 As Integer
    x_soma = 0
    For i = 1 To 9
        x_soma = x_soma + (Val(Mid(x_cpf, i, 1)) * (11 - i))
    Next
    x_soma = (x_soma - Val(x_soma / 11) * 11)
    If x_soma = 1 Or x_soma = 0 Then
        x_digito_1 = 0
    Else
        x_digito_1 = 11 - x_soma
    End If
    x_soma = 0
    For i = 1 To 9
        x_soma = x_soma + (Val(Mid(x_cpf, i, 1)) * (12 - i))
    Next
    x_soma = x_soma + (x_digito_1 * 2)
    x_soma = (x_soma - Val(x_soma / 11) * 11)
    If x_soma = 1 Or x_soma = 0 Then
        x_digito_2 = 0
    Else
        x_digito_2 = 11 - x_soma
    End If
    If Mid(x_cpf, 10, 1) = x_digito_1 And Mid(x_cpf, 11, 1) = x_digito_2 Then
        CalculaDigitoCPF = True
    Else
        CalculaDigitoCPF = False
        MsgBox "O CPF: " & Mid(x_cpf, 1, 3) & "." & Mid(x_cpf, 4, 3) & "." & Mid(x_cpf, 7, 3) & "-" & Mid(x_cpf, 10, 2) & " está com o dígito errado!" & Chr(10) & "O dígito correto é: " & Chr(34) & x_digito_1 & x_digito_2 & Chr(34), vbInformation, "Dígito de CPF Inválido!"
    End If
End Function
Function fContemLetra(x_string As String) As Boolean
    Dim i As Integer
    fContemLetra = False
    i = Len(x_string)
    If i = 0 Then
        Exit Function
    End If
    For i = 1 To Len(x_string)
        If Mid(x_string, i, 1) >= "A" And Mid(x_string, i, 1) <= "Z" Or Mid(x_string, i, 1) >= "a" And Mid(x_string, i, 1) <= "z" Then
            fContemLetra = True
            Exit For
        End If
    Next
End Function
Sub fConverteData()
    If tbl_cheque_baixado.RecordCount > 0 Then
        tbl_cheque_baixado.MoveFirst
        Do Until tbl_cheque_baixado.EOF
            tbl_cheque_baixado.Edit
            tbl_cheque_baixado!Data = tbl_cheque_baixado!Data
            tbl_cheque_baixado.Update
            tbl_cheque_baixado.MoveNext
        Loop
    End If
End Sub
Sub Imprime(f_coluna As Currency, f_linha As Currency, f_string As String, f_local As Integer)
    If f_local = 0 Then
        frm_preview.CurrentX = f_coluna '/ 2
        frm_preview.CurrentY = f_linha '/ 2
        frm_preview.Print f_string
    Else
        Printer.CurrentX = f_coluna
        Printer.CurrentY = f_linha
        Printer.Print f_string
    End If
End Sub
Sub centerform(AnyForm As Form)
    If AnyForm.WindowState = 0 Then
        AnyForm.Top = (Screen.Height - AnyForm.Height) / 2
        AnyForm.Left = (Screen.Width - AnyForm.Width) / 2
    End If
End Sub
Function ImprimeCentralizadoB(ByVal f_string As String, ByVal f_coluna_i As Currency, ByVal f_coluna_f As Currency, ByVal f_linha As Currency, ByVal f_local As Integer) As String
    Dim x_string As String
    x_string = f_string
    If Val(x_string) = 0 Then
        x_string = ""
    End If
    ImprimeCentralizadoB = x_string
    f_coluna_i = f_coluna_i + (f_coluna_f - f_coluna_i - Printer.TextWidth(x_string)) / 2
    Printer.CurrentX = f_coluna_i
    Printer.CurrentY = f_linha
    Imprime f_coluna_i, f_linha, ImprimeCentralizadoB, f_local
End Function
Function ImprimeString(ByVal f_string As String, ByVal f_coluna As Currency, ByVal f_linha As Currency, ByVal f_local As Integer) As String
    ImprimeString = f_string
    Printer.CurrentX = f_coluna
    Printer.CurrentY = f_linha
    Imprime f_coluna, f_linha, ImprimeString, f_local
End Function
Function fTiraCodigoProduto5(ByVal f_codigo As String, ByVal f_tipo As String) As Long
    fTiraCodigoProduto5 = 0
    If Len(f_codigo) >= 4 Then
        If Len(f_codigo) = 4 Then
            f_codigo = "0" & f_codigo
        End If
        If f_tipo = "g" Then
            fTiraCodigoProduto5 = Val(Mid(f_codigo, 1, 2))
        ElseIf f_tipo = "c" Then
            fTiraCodigoProduto5 = Val(Mid(f_codigo, 3, 3))
        End If
    End If
End Function
Function ImprimeCentralizado(ByVal f_string As String, ByVal f_coluna_i As Currency, ByVal f_coluna_f As Currency, ByVal f_linha As Currency, ByVal f_local As Integer) As String
    ImprimeCentralizado = f_string
    f_coluna_i = f_coluna_i + (f_coluna_f - f_coluna_i - Printer.TextWidth(f_string)) / 2
    Printer.CurrentX = f_coluna_i
    Printer.CurrentY = f_linha
    Imprime f_coluna_i, f_linha, ImprimeCentralizado, f_local
End Function
Function ImprimeValorB(ByVal f_valor As String, ByVal f_coluna_i As Currency, ByVal f_coluna_f As Currency, ByVal f_linha As Currency, ByVal f_local As Integer) As String
    Dim x_string As String
    x_string = f_valor
    If Val(x_string) = 0 Then
        x_string = ""
    End If
    ImprimeValorB = x_string
    f_coluna_i = f_coluna_i + (f_coluna_f - f_coluna_i - Printer.TextWidth(x_string) - 0.08)
    Printer.CurrentX = f_coluna_i
    Printer.CurrentY = f_linha
    Imprime f_coluna_i, f_linha, ImprimeValorB, f_local
End Function
Function fValidaValor(valor_x As String) As Currency
    Dim x_posicao As Integer
    Dim x_tamanho As Integer
    Dim x_flag As Boolean
    x_flag = False
    If valor_x <> "" Then
        x_tamanho = Len(valor_x)
        x_posicao = Len(valor_x)
        Do Until x_posicao = 0
            If Not IsNumeric(Mid(valor_x, x_posicao, 1)) Then
                If x_flag = False Then
                    Mid(valor_x, x_posicao, 1) = "."
                    x_flag = True
                Else
                    valor_x = " " & Mid(valor_x, 1, x_posicao - 1) & Mid(valor_x, x_posicao + 1, x_tamanho - x_posicao)
                End If
            End If
            x_posicao = x_posicao - 1
        Loop
    End If
    fValidaValor = Val(Trim(valor_x))
End Function
Function fValidaValor1(valor_x As String) As Currency
    Dim i As Integer
    If valor_x <> "" Then
        i = Len(Trim(valor_x)) - 1
        Mid(valor_x, i, 1) = "."
    End If
    If Len(Trim(valor_x)) = 7 Then
        Mid(valor_x, 1, 2) = " " & Mid(valor_x, 1, 1)
    ElseIf Len(Trim(valor_x)) = 8 Then
        Mid(valor_x, 1, 3) = " " & Mid(valor_x, 1, 2)
    ElseIf Len(Trim(valor_x)) = 9 Then
        Mid(valor_x, 1, 4) = " " & Mid(valor_x, 1, 3)
    End If
    fValidaValor1 = Val(Trim(valor_x))
End Function
Function fValidaValor2(valor_x As String) As Currency
    Dim i As Integer
    If valor_x <> "" Then
        i = Len(Trim(valor_x)) - 2
        Mid(valor_x, i, 1) = "."
    End If
    If Len(Trim(valor_x)) = 8 Then
        Mid(valor_x, 1, 2) = " " & Mid(valor_x, 1, 1)
    ElseIf Len(Trim(valor_x)) = 9 Then
        Mid(valor_x, 1, 3) = " " & Mid(valor_x, 1, 2)
    ElseIf Len(Trim(valor_x)) = 10 Then
        Mid(valor_x, 1, 4) = " " & Mid(valor_x, 1, 3)
    End If
    fValidaValor2 = Val(Trim(valor_x))
End Function
Function fValidaValor4(valor_x As String) As Currency
    Dim i As Integer
    If valor_x <> "" Then
        i = Len(Trim(valor_x)) - 4
        Mid(valor_x, i, 1) = "."
    End If
    If Len(Trim(valor_x)) = 10 Then
        Mid(valor_x, 1, 2) = " " & Mid(valor_x, 1, 1)
    ElseIf Len(Trim(valor_x)) = 11 Then
        Mid(valor_x, 1, 3) = " " & Mid(valor_x, 1, 2)
    ElseIf Len(Trim(valor_x)) = 12 Then
        Mid(valor_x, 1, 4) = " " & Mid(valor_x, 1, 3)
    End If
    fValidaValor4 = Val(Trim(valor_x))
End Function
Function fMascaraTelefone(telefone_x As String) As String
    If Len(Trim(telefone_x)) = 7 Then
        fMascaraTelefone = Mid(telefone_x, 1, 3) & "-" & Mid(telefone_x, 4, 4)
    ElseIf Len(Trim(telefone_x)) = 10 Then
        fMascaraTelefone = "(" & Mid(telefone_x, 1, 3) & ") " & Mid(telefone_x, 4, 3) & "-" & Mid(telefone_x, 7, 4)
    ElseIf Len(Trim(telefone_x)) = 11 Then
        fMascaraTelefone = "(" & Mid(telefone_x, 1, 4) & ") " & Mid(telefone_x, 5, 3) & "-" & Mid(telefone_x, 8, 4)
    Else
        fMascaraTelefone = telefone_x
    End If
End Function
Function fDesmascaraTelefone(telefone_x As String) As String
    If Len(Trim(telefone_x)) = 8 Then
        fDesmascaraTelefone = Mid(telefone_x, 1, 3) & Mid(telefone_x, 5, 4)
    ElseIf Len(Trim(telefone_x)) = 14 Then
        fDesmascaraTelefone = Mid(telefone_x, 2, 3) & Mid(telefone_x, 7, 3) & Mid(telefone_x, 11, 4)
    ElseIf Len(Trim(telefone_x)) = 15 Then
        fDesmascaraTelefone = Mid(telefone_x, 2, 4) & Mid(telefone_x, 8, 3) & Mid(telefone_x, 12, 4)
    Else
        fDesmascaraTelefone = telefone_x
    End If
End Function
Function ImprimeValor(ByVal f_valor As String, ByVal f_coluna_i As Currency, ByVal f_coluna_f As Currency, ByVal f_linha As Currency, ByVal f_local As Integer) As String
    ImprimeValor = f_valor
    f_coluna_i = f_coluna_i + (f_coluna_f - f_coluna_i - Printer.TextWidth(f_valor) - 0.08)
    Printer.CurrentX = f_coluna_i
    Printer.CurrentY = f_linha
    Imprime f_coluna_i, f_linha, ImprimeValor, f_local
End Function
Function ImprimeTexto(ByVal f_string As String, ByVal f_coluna_i As Currency, ByVal f_coluna_f As Currency, ByVal f_linha As Currency, ByVal f_local As Integer) As String
    Dim i As Integer
    ImprimeTexto = ""
    f_coluna_i = f_coluna_i + 0.08
    f_coluna_f = f_coluna_f - 0.08
    Do
        If Trim(f_string) = "" Then
            Exit Do
        End If
        i = i + 1
        If Printer.TextWidth(Mid(f_string, 1, i)) >= (f_coluna_f - f_coluna_i) Then
            Exit Do
        End If
        If Printer.TextWidth(f_string) = 0 Then
            Exit Do
        End If
        ImprimeTexto = Mid(f_string, 1, i)
        If Len(f_string) = i Then
            Exit Do
        End If
    Loop
    Printer.CurrentX = f_coluna_i
    Printer.CurrentY = f_linha
    Imprime f_coluna_i, f_linha, ImprimeTexto, f_local
End Function
Sub ConverteTabela1()
    Dim i As Integer
    Set tbl_programa = bd_sgp.OpenTable("programas")
    tbl_programa.Index = "id_menu"
    tbl_programa.MoveFirst
    Do Until tbl_programa.EOF
        i = i + 1
        tbl_programa.Edit
        tbl_programa!Codigo = i
        tbl_programa.Update
        tbl_programa.MoveNext
    Loop
    tbl_programa.Close
End Sub
Sub GambiConverteMovBombas()
    Dim Obd_sgp As Database
    Dim Nbd_sgp As Database
    Dim Otbl_mov_bomba As Table
    Dim Ntbl_mov_bomba As Table
    
    ChDir "\VB4\SGP"
    Set Obd_sgp = OpenDatabase("SGP_OLD.MDB")
    Set Nbd_sgp = OpenDatabase("SGP_DATA.MDB")
    Set Otbl_mov_bomba = Obd_sgp.OpenTable("movimento_bombas")
    Set Ntbl_mov_bomba = Nbd_sgp.OpenTable("movimento_bombas")
    Otbl_mov_bomba.Index = "id_data"
    Ntbl_mov_bomba.Index = "id_data"
    
    'deleta movimento novo
    Ntbl_mov_bomba.Seek ">=", 4, CDate("01/01/96"), 1, 1
    Do Until Ntbl_mov_bomba.EOF
        If Ntbl_mov_bomba!Empresa <> 4 Then
            Exit Do
        End If
        Ntbl_mov_bomba.Edit
        Ntbl_mov_bomba.Delete
        Ntbl_mov_bomba.MoveNext
    Loop
    
    'lê movimento antigo
    Otbl_mov_bomba.Seek ">=", 4, CDate("01/01/96"), 1, 1
    Do Until Otbl_mov_bomba.EOF
        If Otbl_mov_bomba!Empresa <> 4 Then
            Exit Do
        End If
        Ntbl_mov_bomba.AddNew
        Ntbl_mov_bomba!Empresa = Otbl_mov_bomba!Empresa
        Ntbl_mov_bomba!Data = Otbl_mov_bomba!Data
        Ntbl_mov_bomba!Periodo = Otbl_mov_bomba!Periodo
        Ntbl_mov_bomba!bomba = Otbl_mov_bomba!bomba
        Ntbl_mov_bomba!Abertura = Otbl_mov_bomba!Abertura
        Ntbl_mov_bomba!Encerrante = Otbl_mov_bomba!Encerrante
        Ntbl_mov_bomba!saida = Otbl_mov_bomba!saida
        Ntbl_mov_bomba!custo = Otbl_mov_bomba!custo
        Ntbl_mov_bomba!venda = Otbl_mov_bomba!venda
        Ntbl_mov_bomba!tipo_combustivel = Otbl_mov_bomba!tipo_combustivel
        Ntbl_mov_bomba!tanque = Otbl_mov_bomba!tanque
        Ntbl_mov_bomba.Update
        Otbl_mov_bomba.MoveNext
    Loop
End Sub
Sub ErroArquivo(x_arquivo As String, x_nome As String)
    x_arquivo = UCase(x_arquivo)
    If Err = 3022 Then
        MsgBox Mid(x_nome, 1, Len(x_nome) - 1) & " já Cadastrad" & Mid(x_nome, Len(x_nome), 1) & "!" & Chr(10) & Error, 48, "Erro na Tabela: " & x_arquivo
        Exit Sub
    ElseIf Err = 3019 Then
        MsgBox "Índice não informado!" & Chr(10) & Error, 48, "Erro na Tabela: " & x_arquivo
        Exit Sub
    ElseIf Err = 3058 Then
        MsgBox "Campo do índice não informado!" & Chr(10) & Error, 48, "Erro na Tabela: " & x_arquivo
        Exit Sub
    ElseIf Err = 3265 Then
        MsgBox "Variável inexistente!" & Chr(10) & Error, 48, "Erro na Tabela: " & x_arquivo
        Exit Sub
    ElseIf Err = 3315 Then
        MsgBox "Campo com valor nulo!" & Chr(10) & Error, 48, "Erro na Tabela: " & x_arquivo
        Exit Sub
    Else
        MsgBox "Erro de número " & Err & Chr(10) & Error, 48, "Erro na Tabela: " & x_arquivo
        Exit Sub
    End If
End Sub
Sub BaixaVencimentos(data_x As Date)
    On Error GoTo FileError
    Dim x_empresa As Integer
    Dim x_data As Date
    menu_personalizado2.dta_empresa.Recordset.MoveFirst
    If menu_personalizado2.dta_empresa.Recordset.Nome = "Auto Posto Nova Suíça Ltda" Then
        If (MsgBox("O sistema irá dar baixa automática nos vencimentos antes da data " & Chr(34) & data_x & Chr(34) & "." & Chr(10) & Chr(10) & " Confirme para continuar.", 4 + 32 + 0, "Baixa Automática dos Vencimentos!")) = 6 Then
            Set tbl_contas_pagar = bd_sgp.OpenTable("contas_pagar")
            Set tbl_baixa_pagar = bd_sgp.OpenTable("baixa_pagar")
            tbl_contas_pagar.Index = "id_data_vencimento"
            tbl_baixa_pagar.Index = "id_registro"
            For x_empresa = 1 To 11
                x_data = "01/01/80"
                tbl_contas_pagar.Seek ">", x_empresa, x_data, "          ", 0
                If Not tbl_contas_pagar.NoMatch Then
                    Do Until tbl_contas_pagar.EOF
                        If tbl_contas_pagar!Empresa <> x_empresa Then
                            Exit Do
                        End If
                        If Not tbl_contas_pagar!Data_Vencimento < data_x Then
                            Exit Do
                        End If
                        tbl_baixa_pagar.Seek "=", tbl_contas_pagar!Registro
                        If tbl_baixa_pagar.NoMatch Then
                            tbl_contas_pagar.Edit
                            tbl_baixa_pagar.AddNew
                            tbl_baixa_pagar!Empresa = tbl_contas_pagar!Empresa
                            tbl_baixa_pagar!Registro = tbl_contas_pagar!Registro
                            tbl_baixa_pagar!codigo_fornecedor = tbl_contas_pagar!codigo_fornecedor
                            tbl_baixa_pagar!nome_fornecedor = tbl_contas_pagar!nome_fornecedor
                            tbl_baixa_pagar!data_emissao = tbl_contas_pagar!data_emissao
                            tbl_baixa_pagar!Data_Vencimento = tbl_contas_pagar!Data_Vencimento
                            tbl_baixa_pagar!Valor = tbl_contas_pagar!Valor
                            tbl_baixa_pagar!Numero_Documento = tbl_contas_pagar!Numero_Documento
                            tbl_baixa_pagar!local_cobranca = tbl_contas_pagar!local_cobranca
                            tbl_baixa_pagar!codigo_conta = tbl_contas_pagar!codigo_conta
                            tbl_baixa_pagar!Complemento = tbl_contas_pagar!Complemento
                            tbl_baixa_pagar![Data da Digitacao] = tbl_contas_pagar![Data da Digitacao]
                            tbl_baixa_pagar!data_pagamento = tbl_contas_pagar!Data_Vencimento
                            tbl_baixa_pagar!valor_pagamento = tbl_contas_pagar!Valor
                            tbl_baixa_pagar.Update
                            tbl_contas_pagar.Delete
                        End If
                        tbl_contas_pagar.MoveNext
                    Loop
                End If
            Next
            tbl_contas_pagar.Close
            tbl_baixa_pagar.Close
        End If
    End If
    Exit Sub
FileError:
    ErroArquivo tbl_baixa_pagar.Name, "Baixa de Vencimentoa"
    Exit Sub
End Sub
Sub BaixaChequePreDatado(x_data As Date)
    On Error GoTo FileError
    Dim x_empresa As Integer
    Dim tbl_baixa_cheque As Table
    Dim tbl_movimento_cheque As Table
    menu_personalizado2.dta_empresa.Recordset.MoveFirst
    If menu_personalizado2.dta_empresa.Recordset.Nome = "Auto Posto Nova Suíça Ltda" Then
        If (MsgBox("O sistema irá dar baixa automática nos cheques antes da data " & Chr(34) & Format(x_data, "dd/mm/yyyy") & Chr(34) & "." & Chr(10) & Chr(10) & " Confirme para continuar.", 4 + 32 + 256, "Baixa Automática dos Cheques Pré-Datados!")) = 6 Then
            Set tbl_baixa_cheque = bd_sgp_b.OpenTable("Baixa_Cheque")
            Set tbl_movimento_cheque = bd_sgp_m.OpenTable("Movimento_Cheque")
            tbl_baixa_cheque.Index = "id_vencimento"
            tbl_movimento_cheque.Index = "id_vencimento"
            With tbl_movimento_cheque
                If .RecordCount > 0 Then
                    For x_empresa = 1 To 11
                        .Seek "<=", x_empresa, CDate(x_data), CDate("01/01/1900"), "1", 0, "          ", "      "
                        If Not .NoMatch Then
                            Do Until .BOF
                                If !Empresa <> x_empresa Or ![Data do Vencimento] >= x_data Then
                                    Exit Do
                                End If
                                .Edit
                                tbl_baixa_cheque.AddNew
                                tbl_baixa_cheque!Empresa = !Empresa
                                tbl_baixa_cheque![Data de Emissao] = ![Data de Emissao]
                                tbl_baixa_cheque![Numero da Conta] = ![Numero da Conta]
                                tbl_baixa_cheque![Numero do Cheque] = ![Numero do Cheque]
                                tbl_baixa_cheque!Periodo = !Periodo
                                tbl_baixa_cheque![Tipo do Movimento] = ![Tipo do Movimento]
                                tbl_baixa_cheque!Valor = !Valor
                                tbl_baixa_cheque![Data do Vencimento] = ![Data do Vencimento]
                                tbl_baixa_cheque!Emitente = !Emitente
                                tbl_baixa_cheque![Ordem da Digitacao] = ![Ordem da Digitacao]
                                tbl_baixa_cheque![Codigo de Barra 1] = ![Codigo de Barra 1]
                                tbl_baixa_cheque![Codigo de Barra 2] = ![Codigo de Barra 2]
                                tbl_baixa_cheque![Codigo de Barra 3] = ![Codigo de Barra 3]
                                tbl_baixa_cheque![Banco Agencia] = ![Banco Agencia]
                                tbl_baixa_cheque.Update
                                .Delete
                                .MovePrevious
                            Loop
                        End If
                    Next
                    MsgBox "Baixa de cheques concluida!", 48, "Baixa de Cheque Pré-Datado."
                End If
            End With
            tbl_baixa_cheque.Close
            tbl_movimento_cheque.Close
        End If
    End If
    Exit Sub
FileError:
    ErroArquivo tbl_baixa_cheque.Name, "Cheque Baixadoo"
    Exit Sub
End Sub
Sub CentraForm(unForm As Form)
    unForm.Left = (Screen.Width - unForm.Width) / 2
    unForm.Top = (Screen.Height - unForm.Height) / 2
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
Function FazExtenso(ByVal valor_x As Currency) As String
    FazExtenso = ""
    'Esta rotina faz o extenso até o limite de
    '900.000,00 (Novecentos Mil reais)
    Dim x_extenso As String
    Dim Valor As String
    'Cria variáveis
    Dim unidade(1 To 9) As String
    Dim dezena_1(0 To 9) As String
    Dim dezena(2 To 9) As String
    Dim centena(1 To 9) As String
    'unidade
    unidade(1) = "Um "
    unidade(2) = "Dois "
    unidade(3) = "Três "
    unidade(4) = "Quatro "
    unidade(5) = "Cinco "
    unidade(6) = "Seis "
    unidade(7) = "Sete "
    unidade(8) = "Oito "
    unidade(9) = "Nove "
    'dezena_1
    dezena_1(0) = "Dez "
    dezena_1(1) = "Onze "
    dezena_1(2) = "Doze "
    dezena_1(3) = "Treze "
    dezena_1(4) = "Quatrorze "
    dezena_1(5) = "Quinze "
    dezena_1(6) = "Dezesseis "
    dezena_1(7) = "Dezesete "
    dezena_1(8) = "Dezoito "
    dezena_1(9) = "Dezenove "
    'dezena
    dezena(2) = "Vinte "
    dezena(3) = "Trinta "
    dezena(4) = "Quarenta "
    dezena(5) = "Cinquenta "
    dezena(6) = "Sessenta "
    dezena(7) = "Setenta "
    dezena(8) = "Oitenta "
    dezena(9) = "Noventa "
    'centena
    centena(1) = "Cento "
    centena(2) = "Duzentos "
    centena(3) = "Trezentos "
    centena(4) = "Quatrocentos "
    centena(5) = "Quinhentos "
    centena(6) = "Seiscentos "
    centena(7) = "Setecentos "
    centena(8) = "Oitocentos "
    centena(9) = "Novecentos "
    
    x_extenso = ""
    Valor = Format(valor_x, "0000000000.00")
    
    '*** início (1000 a 999000) ***
    'testa centena de milhar de (100 000 a 900 000)
    If Mid(Valor, 5, 1) > 0 Then
        If Mid(Valor, 5, 3) = 100 Then
            x_extenso = x_extenso & "Cem Mil "
        Else
            x_extenso = x_extenso & centena(Mid(Valor, 5, 1))
            If Mid(Valor, 6, 2) > 0 Then
                x_extenso = x_extenso & "e "
            End If
        End If
    End If
    'testa dezena de milhar de (10 a 19)
    If Mid(Valor, 6, 1) = 1 Then
        x_extenso = x_extenso & dezena_1(Mid(Valor, 7, 1)) & "Mil "
    End If
    'testa dezena de milhar (20 a 90)
    If Mid(Valor, 6, 1) > 1 Then
        x_extenso = x_extenso & dezena(Mid(Valor, 6, 1))
        If Mid(Valor, 7, 1) > 0 Then
            x_extenso = x_extenso & "e "
        Else
            x_extenso = x_extenso & "Mil "
        End If
    End If
    'testa unidade de milhar (1 a 9) e dezena de milhar <> 1
    If Mid(Valor, 7, 1) > 0 And Mid(Valor, 6, 1) <> 1 Then
        x_extenso = x_extenso & unidade(Mid(Valor, 7, 1)) & "Mil "
    End If
    'testa de milhar(xxx xxx) e > 0 e unidade(xxx) > 0
    If Mid(Valor, 5, 3) > 0 And Mid(Valor, 8, 3) > 0 Then
        x_extenso = x_extenso & "e "
    End If
    '*** fim (1000 a 999000) ***
    
    
    '*** início (1 a 999) ***
    'testa centena de (100 a 900)
    If Mid(Valor, 8, 1) > 0 Then
        If Mid(Valor, 8, 3) = 100 Then
            x_extenso = x_extenso & "Cem "
        Else
            x_extenso = x_extenso & centena(Mid(Valor, 8, 1))
            If Mid(Valor, 9, 2) > 0 Then
                x_extenso = x_extenso & "e "
            End If
        End If
    End If
    'testa dezena de (10 a 19)
    If Mid(Valor, 9, 1) = 1 Then
        x_extenso = x_extenso & dezena_1(Mid(Valor, 10, 1))
    End If
    'testa dezena de (20 a 90)
    If Mid(Valor, 9, 1) > 1 Then
        x_extenso = x_extenso & dezena(Mid(Valor, 9, 1))
        If Mid(Valor, 10, 1) > 0 Then
            x_extenso = x_extenso & "e "
        End If
    End If
    'testa unidade de (1 a 9) e dezena <> 1
    If Mid(Valor, 10, 1) > 0 And Mid(Valor, 9, 1) <> 1 Then
        x_extenso = x_extenso & unidade(Mid(Valor, 10, 1))
    End If
    '*** fim (1 a 999) ***
    
    
    'testa se é Real ou Reais
    If Mid(Valor, 1, 10) = 1 Then
        x_extenso = x_extenso & "Real "
    ElseIf Mid(Valor, 1, 10) > 1 Then
        x_extenso = x_extenso & "Reais "
    End If
    
    
    '*** início Centavos (1 a 99) ***
    'testa dezena de (10 a 19)
    If Mid(Valor, 12, 2) > 0 Then
        x_extenso = x_extenso & "e "
    End If
    'testa dezena de centavos de (10 a 19)
    If Mid(Valor, 12, 1) = 1 Then
        x_extenso = x_extenso & dezena_1(Mid(Valor, 13, 1))
    End If
    'testa dezena de centavos de (20 a 90)
    If Mid(Valor, 12, 1) > 1 Then
        x_extenso = x_extenso & dezena(Mid(Valor, 12, 1))
        If Mid(Valor, 13, 1) > 0 Then
            x_extenso = x_extenso & "e "
        End If
    End If
    'testa unidade de centavos de (1 a 9) e dezena de centavos <> 1
    If Mid(Valor, 13, 1) > 0 And Mid(Valor, 12, 1) <> 1 Then
        x_extenso = x_extenso & unidade(Mid(Valor, 13, 1))
    End If
    '*** fim Centavos (1 a 99) ***
    
    
    'testa se é Centavo ou Centavos
    If Mid(Valor, 12, 2) = 1 Then
        x_extenso = x_extenso & "Centavo "
    ElseIf Mid(Valor, 12, 2) > 1 Then
        x_extenso = x_extenso & "Centavos "
    End If
    
    Mid(x_extenso, Len(Trim(x_extenso)) + 1, 1) = "."
    FazExtenso = x_extenso
End Function
Sub FazExtensoMes(data_x As String, mes_x As String)
    Select Case Month(data_x)
        Case 1
            mes_x = "Janeiro  "
        Case 2
            mes_x = "Fevereiro"
        Case 3
            mes_x = "Março    "
        Case 4
            mes_x = "Abril    "
        Case 5
            mes_x = "Maio     "
        Case 6
            mes_x = "Junho    "
        Case 7
            mes_x = "Julho    "
        Case 8
            mes_x = "Agosto   "
        Case 9
            mes_x = "Setembro "
        Case 10
            mes_x = "Outubro  "
        Case 11
            mes_x = "Novembro "
        Case 12
            mes_x = "Dezembro "
    End Select
End Sub
Sub PaiFilho(pai_x As String, filho_x As String)
    Dim x_tamanho As Integer
    Dim x_posicao As Integer
    Dim x_posicao2 As Integer
    Dim x_pai As String
    Dim x_filho As String
    If pai_x = "\" Then
        pai_x = "\"
        filho_x = "\"
        Exit Sub
    End If
    x_tamanho = Len(Trim(pai_x))
    x_posicao = x_tamanho
    Do Until x_posicao = 0
        If Mid(pai_x, x_posicao, 1) = "\" Then
            x_filho = Mid(pai_x, (x_posicao + 1), (x_tamanho - x_posicao))
            Exit Do
        End If
        x_posicao = x_posicao - 1
    Loop
    x_tamanho = x_tamanho - (Len(Trim(x_filho)) + 1)
    x_posicao = x_tamanho
    Do Until x_posicao = 0
        If Mid(pai_x, x_posicao, 1) = "\" Then
            x_pai = Mid(pai_x, (x_posicao + 1), (x_tamanho - x_posicao))
            Exit Do
        End If
        x_posicao = x_posicao - 1
    Loop
    If pai_x = "" Then
        pai_x = "\"
    End If
    pai_x = x_pai
    filho_x = x_filho
End Sub
Function RetiraGString(numero As Integer) As String
    RetiraGString = ""
    Dim x_index As Integer
    Dim x_inicio As Integer
    Dim x_numero As Integer
    x_inicio = 1
    x_numero = 1
    If Len(g_string) > 0 Then
        Do Until x_index > Len(g_string)
            x_index = x_index + 1
            If Mid(g_string, x_index, 3) = "|@|" Then
                If x_numero = numero Then
                    RetiraGString = Mid(g_string, x_inicio, x_index - x_inicio)
                    Exit Function
                End If
                x_index = x_index + 2
                x_numero = x_numero + 1
                x_inicio = x_index + 1
            End If
        Loop
    End If
End Function
Sub CriaCampo(x_nome_campo As String, x_tipo_campo As String, x_tamanho_campo As Integer, x_campo_em_branco As String)
    Set campo = New Field
    campo.Name = Trim(x_nome_campo)
    If Trim(x_tipo_campo) = "Data/Time" Then
        campo.Type = dbDate
    ElseIf Trim(x_tipo_campo) = "Text" Then
        campo.Type = dbText
    ElseIf Trim(x_tipo_campo) = "Memo" Then
        campo.Type = dbMemo
    ElseIf Trim(x_tipo_campo) = "Boolean" Then
        campo.Type = dbBoolean
    ElseIf Trim(x_tipo_campo) = "Integer" Then
        campo.Type = dbInteger
    ElseIf Trim(x_tipo_campo) = "Long" Then
        campo.Type = dbLong
    ElseIf Trim(x_tipo_campo) = "Currency" Then
        campo.Type = dbCurrency
    ElseIf Trim(x_tipo_campo) = "Single" Then
        campo.Type = dbSingle
    ElseIf Trim(x_tipo_campo) = "Double" Then
        campo.Type = dbDouble
    ElseIf Trim(x_tipo_campo) = "Byte" Then
        campo.Type = dbByte
    ElseIf Trim(x_tipo_campo) = "Long Binary" Then
        campo.Type = dbLongBinary
    End If
    If x_tamanho_campo > 0 Then
        campo.Size = x_tamanho_campo
    End If
    If Trim(x_tipo_campo) = "Text" Or Trim(x_tipo_campo) = "Memo" Then
        If x_campo_em_branco = "S" Then
            campo.AllowZeroLength = True
        End If
    End If
    tabela.Fields.Append campo
End Sub
Sub CriaIndice(x_nome_indice As String, x_campo_do_indice As String, x_indice_primario As String, x_indice_unico As String)
    Set indice = New Index
    indice.Name = Trim(x_nome_indice)
    indice.Fields = Trim(x_campo_do_indice)
    If Trim(x_indice_primario) = "S" Then
        indice.Primary = True
    Else
        indice.Primary = False
    End If
    If Trim(x_indice_unico) = "S" Then
        indice.Unique = True
    Else
        indice.Unique = False
    End If
    tabela.Indexes.Append indice
End Sub
Sub CriaTabelaEstoque()
    Set tabela = New TableDef
    tabela.Name = "Estoque"
    CriaCampo "Empresa", "Integer", 0, ""
    CriaCampo "Grupo do Produto", "Integer", 0, ""
    CriaCampo "Codigo do Produto", "Long", 0, ""
    CriaCampo "Codigo do Produto2", "Long", 0, ""
    CriaCampo "Quantidade", "Currency", 0, "N"
    CriaCampo "Preco de Venda", "Currency", 0, "N"
    CriaIndice "id_codigo", "Empresa;Grupo do Produto;Codigo do Produto", "S", "S"
    CriaIndice "id_codigo2", "Empresa;Codigo do Produto2", "N", "S"
    bd_disco.TableDefs.Append tabela
End Sub
Sub CriaTabelaGrupo()
    Set tabela = New TableDef
    tabela.Name = "Grupo"
    CriaCampo "Codigo", "Integer", 0, ""
    CriaCampo "Nome", "Text", 35, "N"
    CriaIndice "id_codigo", "Codigo", "S", "S"
    CriaIndice "id_nome", "Nome;Codigo", "N", "S"
    bd_disco.TableDefs.Append tabela
End Sub
Sub CriaTabelaProduto()
    Set tabela = New TableDef
    tabela.Name = "Produto"
    CriaCampo "Codigo do Grupo", "Integer", 0, ""
    CriaCampo "Codigo do Produto", "Long", 0, ""
    CriaCampo "Digito", "Integer", 0, ""
    CriaCampo "Codigo2", "Long", 0, ""
    CriaCampo "Nome", "Text", 40, "N"
    CriaCampo "Unidade", "Text", 3, "N"
    CriaCampo "Preco de Venda", "Currency", 0, "N"
    CriaCampo "Preco de Custo Anterior", "Currency", 0, "N"
    CriaCampo "Preco de Custo", "Currency", 0, "N"
    CriaCampo "Preco de Custo Medio", "Currency", 0, "N"
    CriaCampo "Estoque Minimo", "Currency", 0, "N"
    CriaCampo "Estoque Maximo", "Currency", 0, "N"
    CriaCampo "Estoque Fisico", "Currency", 0, "N"
    CriaCampo "Data da Ultima Compra", "Data/Time", 0, "N"
    CriaCampo "Data do Acerto", "Data/Time", 0, "N"
    CriaCampo "Comissao", "Boolean", 0, "N"
    CriaCampo "Inativo", "Boolean", 0, "N"
    CriaCampo "Codigo da Aliquota", "Integer", 0, ""
    CriaIndice "id_codigo", "Codigo do Grupo;Codigo do Produto", "S", "S"
    CriaIndice "id_codigo2", "Codigo2", "N", "S"
    CriaIndice "id_grupo", "Codigo do Grupo;Nome;Codigo2", "N", "S"
    CriaIndice "id_nome", "Nome;Codigo do Grupo;Codigo do Produto", "N", "S"
    bd_disco.TableDefs.Append tabela
End Sub
Sub DuplicataReceberVbCobol()
    Dim dados As String
    Dim x_string As String * 40
    Set tbl_duplicata_receber = bd_sgp.OpenTable("Duplicata_Receber")
    Open "\CONVERTE\ARQ\SCR-DUPL.TXT" For Output As #1
    With tbl_duplicata_receber
        .Index = "id_documento"
        .MoveFirst
        Do Until .EOF
            dados = Format(!Empresa, "00")
            dados = dados & Format(![Numero do Documento], "0000000000")
            dados = dados & Format(![Codigo do Cliente], "000000")
            dados = dados & Format(![Data de Emissao], "yymmdd")
            dados = dados & Format(![Data do Periodo Inicial], "yymmdd")
            dados = dados & Format(![Data do Periodo Final], "yymmdd")
            dados = dados & Format(![Numero da Duplicata], "00000000000")
            dados = dados & Format(![Data do Vencimento], "yymmdd")
            dados = dados & Mid(Format(![Valor do Custo Bancario], "0000000000.00"), 1, 10) & Mid(Format(![Valor do Custo Bancario], "0000000000.00"), 12, 2)
            dados = dados & Mid(Format(![Valor do Vencimento], "0000000000.00"), 1, 10) & Mid(Format(![Valor do Vencimento], "0000000000.00"), 12, 2)
            Print #1, dados
            tbl_duplicata_receber.MoveNext
        Loop
        Close #1
        .Close
    End With
End Sub
Sub ScpVbCobol()
    Dim dados As String
    Dim x_string As String * 40
    Set tbl_contas_pagar = bd_sgp.OpenTable("contas_pagar")
    Open "\CONVERTE\ARQ\SCP-VENC.TXT" For Output As #1
    tbl_contas_pagar.Index = "id_registro"
    tbl_contas_pagar.MoveFirst
    Do Until tbl_contas_pagar.EOF
        dados = Format(tbl_contas_pagar!Empresa, "00")
        dados = dados & Format(tbl_contas_pagar!Registro, "000000")
        dados = dados & Format(tbl_contas_pagar!codigo_fornecedor, "0000")
        x_string = tbl_contas_pagar!nome_fornecedor
        dados = dados & x_string
        dados = dados & Format(tbl_contas_pagar!data_emissao, "yymmdd")
        dados = dados & Format(tbl_contas_pagar!Data_Vencimento, "yymmdd")
        x_string = Format(tbl_contas_pagar!Valor, "0000000000.00")
        Mid(x_string, 11, 2) = Mid(x_string, 12, 2)
        dados = dados & Mid(x_string, 1, 12)
        x_string = tbl_contas_pagar!Numero_Documento
        dados = dados & Mid(x_string, 1, 20)
        dados = dados & Format(tbl_contas_pagar!local_cobranca, "000")
        dados = dados & Format(tbl_contas_pagar!codigo_conta, "000")
        x_string = tbl_contas_pagar!Complemento
        dados = dados & x_string
        Print #1, dados
        tbl_contas_pagar.MoveNext
    Loop
    Close #1
    tbl_contas_pagar.Close
End Sub

