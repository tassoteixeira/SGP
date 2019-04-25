VERSION 5.00
Begin VB.Form rel_pl_bombas 
   Caption         =   "Emissão das Planilhas de Bmbas"
   ClientHeight    =   1425
   ClientLeft      =   2145
   ClientTop       =   2100
   ClientWidth     =   6285
   Icon            =   "REL_PL_BOMBAS.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "REL_PL_BOMBAS.frx":0442
   ScaleHeight     =   1425
   ScaleWidth      =   6285
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4980
      Picture         =   "REL_PL_BOMBAS.frx":0488
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   360
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3600
      Picture         =   "REL_PL_BOMBAS.frx":1762
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Imprime planilha de bomba."
      Top             =   360
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
      Begin VB.TextBox txt_pagina 
         Height          =   285
         Left            =   1800
         MaxLength       =   3
         TabIndex        =   2
         Text            =   "2"
         Top             =   480
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "&Número de Páginas"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "rel_pl_bombas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_margem_lin As Currency
Dim l_margem_col As Currency
Dim l_lin_max As Currency
Dim l_col_max As Currency
Dim l_local As Integer
Dim l_impressos As Integer
Private Sub ZeraVariaveis()
    Dim i As Integer
    l_margem_col = 0.8
    l_margem_lin = 0.5
    l_col_max = 20
    l_lin_max = 28
End Sub
Private Sub ImprimeGrade()
    l_local = 1
    'Seleciona tamanho da fonte
    Printer.FontSize = 34
    'Printer.Line (0, 0)-(0, 26)
    Printer.DrawWidth = 8
    Printer.Line (l_margem_col, l_margem_lin)-(l_col_max, l_lin_max), , B
    Printer.FontSize = 14
    Printer.DrawWidth = 6
    ImprimeCentralizado "MOVIMENTAÇÃO DAS BOMBAS", l_margem_col, l_col_max, l_margem_lin + 0.1, l_local
    Printer.FontSize = 8
    ImprimeString "Frentista:", l_margem_col + 0.5, l_margem_lin + 1.5, l_local
    ImprimeString "Data: ____/____/____", l_margem_col + 9, l_margem_lin + 1.5, l_local
    ImprimeString "Periodo: ____:____", l_margem_col + 15, l_margem_lin + 1.5, l_local
    Printer.DrawWidth = 3
    Printer.Line (l_margem_col, l_margem_lin + 2)-(l_col_max, l_margem_lin + 2)
    
    Printer.Line (l_margem_col, l_margem_lin + 3.4)-(l_col_max, l_margem_lin + 3.4)
    
    ImprimeCentralizado "ABERTURA", l_margem_col, l_margem_col + 4.5, l_margem_lin + 2.5, l_local
    ImprimeCentralizado "FECHAMENTO", l_margem_col + 4.5, l_margem_col + 9, l_margem_lin + 2.5, l_local
    ImprimeCentralizado "LITROS", l_margem_col + 9, l_margem_col + 12.2, l_margem_lin + 2.5, l_local
    ImprimeCentralizado "PREÇO P/ LITRO", l_margem_col + 12.2, l_margem_col + 15.2, l_margem_lin + 2.5, l_local
    ImprimeCentralizado "TOTAL EM R$", l_margem_col + 15.2, l_col_max, l_margem_lin + 2.5, l_local
    
    Printer.Line (l_margem_col + 4.5, l_margem_lin + 2)-(l_margem_col + 4.5, l_margem_lin + 19.8)
    Printer.Line (l_margem_col + 9, l_margem_lin + 2)-(l_margem_col + 9, l_margem_lin + 19.8)
    Printer.Line (l_margem_col + 12.2, l_margem_lin + 2)-(l_margem_col + 12.2, l_margem_lin + 18.5)
    Printer.Line (l_margem_col + 15.2, l_margem_lin + 2)-(l_margem_col + 15.2, l_margem_lin + 18.5)
    Printer.Line (l_margem_col, l_margem_lin + 4.8)-(l_col_max, l_margem_lin + 4.8)
    
    Printer.Line (l_margem_col, l_margem_lin + 6.1)-(l_col_max, l_margem_lin + 6.1)
    
    Printer.Line (l_margem_col, l_margem_lin + 7.4)-(l_col_max, l_margem_lin + 7.4)
    
    Printer.Line (l_margem_col, l_margem_lin + 8.7)-(l_col_max, l_margem_lin + 8.7)
    
    Printer.Line (l_margem_col, l_margem_lin + 10.1)-(l_col_max, l_margem_lin + 10.1)
    
    Printer.Line (l_margem_col, l_margem_lin + 11.5)-(l_col_max, l_margem_lin + 11.5)
    
    Printer.Line (l_margem_col, l_margem_lin + 12.9)-(l_col_max, l_margem_lin + 12.9)
    
    Printer.Line (l_margem_col, l_margem_lin + 14.3)-(l_col_max, l_margem_lin + 14.3)
    ImprimeCentralizado "GASOLINA", l_margem_col, l_margem_col + 4.5, l_margem_lin + 14.6, l_local
    ImprimeCentralizado "TOTAL", l_margem_col + 9, l_margem_col + 12.2, l_margem_lin + 14.6, l_local
    
    Printer.Line (l_margem_col, l_margem_lin + 15.7)-(l_col_max, l_margem_lin + 15.7)
    ImprimeCentralizado "ÁLCOOL", l_margem_col, l_margem_col + 4.5, l_margem_lin + 16, l_local
    ImprimeCentralizado "CAIXA", l_margem_col + 9, l_margem_col + 12.2, l_margem_lin + 16, l_local
    
    Printer.Line (l_margem_col, l_margem_lin + 17.1)-(l_col_max, l_margem_lin + 17.1)
    ImprimeCentralizado "DIESEL", l_margem_col, l_margem_col + 4.5, l_margem_lin + 17.5, l_local
    ImprimeCentralizado "TOTAL", l_margem_col + 9, l_margem_col + 12.2, l_margem_lin + 17.5, l_local
    
    Printer.Line (l_margem_col, l_margem_lin + 18.5)-(l_col_max, l_margem_lin + 18.5)
    ImprimeCentralizado "TOTAL", l_margem_col, l_margem_col + 4.5, l_margem_lin + 19, l_local
    
    Printer.Line (l_margem_col, l_margem_lin + 19.8)-(l_col_max, l_margem_lin + 19.8)
    Printer.Line (l_margem_col + 5, l_margem_lin + 19.8)-(l_margem_col + 5, l_lin_max)
    Printer.Line (l_margem_col + 10, l_margem_lin + 19.8)-(l_margem_col + 10, l_lin_max)
    
    Printer.Line (l_margem_col, l_margem_lin + 20.9)-(l_col_max, l_margem_lin + 20.9)
    ImprimeCentralizado "CHEQUE PRÉ-DATADO", l_margem_col, l_margem_col + 4.5, l_margem_lin + 20.3, l_local
    ImprimeString "Observação:", l_margem_col + 10.5, l_margem_lin + 20.3, l_local
    
    Printer.Line (l_margem_col, l_margem_lin + 21.8)-(l_col_max, l_margem_lin + 21.8)
    ImprimeCentralizado "CHEQUE VISTA", l_margem_col, l_margem_col + 4.5, l_margem_lin + 21.3, l_local
    
    Printer.Line (l_margem_col, l_margem_lin + 22.8)-(l_col_max, l_margem_lin + 22.8)
    ImprimeCentralizado "DINHEIRO", l_margem_col, l_margem_col + 4.5, l_margem_lin + 22.3, l_local
    
    Printer.Line (l_margem_col, l_margem_lin + 23.7)-(l_col_max, l_margem_lin + 23.7)
    ImprimeCentralizado "NOTA FIRMA", l_margem_col, l_margem_col + 4.5, l_margem_lin + 23.2, l_local
    
    Printer.Line (l_margem_col, l_margem_lin + 24.7)-(l_col_max, l_margem_lin + 24.7)
    ImprimeCentralizado "T O T A L", l_margem_col, l_margem_col + 4.5, l_margem_lin + 25, l_local
    
    Printer.Line (l_margem_col, l_margem_lin + 25.6)-(l_col_max, l_margem_lin + 25.6)
    
    Printer.Line (l_margem_col, l_margem_lin + 26.6)-(l_col_max, l_margem_lin + 26.6)
    
    Printer.Line (l_margem_col, l_margem_lin + 27.5)-(l_col_max, l_margem_lin + 27.5)
    
    'Printer.Line (l_margem_col, l_margem_lin + 1.6)-(l_col_max, l_margem_lin + 1.6)
    'Printer.Line (l_margem_col + 14.8, l_margem_lin + 0.8)-(17, l_margem_lin + 1.6)
    'ImprimeString "1 Produto:", l_margem_col + 0.1, l_margem_lin + 1.05, l_local
    'ImprimeString "2 Data:", l_margem_col + 15, l_margem_lin + 1.05, l_local
    
    'Printer.Line (l_margem_col, l_margem_lin + 2.4)-(l_col_max, l_margem_lin + 2.4)
    'ImprimeString "3 Estoque de abertura (medição física no início do dia)", l_margem_col + 0.1, l_margem_lin + 1.85, l_local
    
    'Printer.Line (l_margem_col, l_margem_lin + 3.2)-(l_col_max, l_margem_lin + 3.2)
    'Printer.Line (l_margem_col + 1.4, l_margem_lin + 2.4)-(l_margem_col + 1.4, l_margem_lin + 3.2)
    'Printer.Line (l_margem_col + 2, l_margem_lin + 2.4)-(l_margem_col + 2, l_margem_lin + 4)
    'Printer.Line (l_margem_col + 3.4, l_margem_lin + 2.4)-(l_margem_col + 3.4, l_margem_lin + 3.2)
    'Printer.Line (l_margem_col + 4, l_margem_lin + 2.4)-(l_margem_col + 4, l_margem_lin + 4)
    'Printer.Line (l_margem_col + 5.4, l_margem_lin + 2.4)-(l_margem_col + 5.4, l_margem_lin + 3.2)
    'Printer.Line (l_margem_col + 6, l_margem_lin + 2.4)-(l_margem_col + 6, l_margem_lin + 4)
    'Printer.Line (l_margem_col + 7.4, l_margem_lin + 2.4)-(l_margem_col + 7.4, l_margem_lin + 3.2)
    'Printer.Line (l_margem_col + 8, l_margem_lin + 2.4)-(l_margem_col + 8, l_margem_lin + 4)
    'Printer.Line (l_margem_col + 9.4, l_margem_lin + 2.4)-(l_margem_col + 9.4, l_margem_lin + 3.2)
    'Printer.Line (l_margem_col + 10, l_margem_lin + 2.4)-(l_margem_col + 10, l_margem_lin + 21.9)
    'Printer.Line (l_margem_col + 11.4, l_margem_lin + 2.4)-(l_margem_col + 11.4, l_margem_lin + 3.2)
    'Printer.Line (l_margem_col + 12, l_margem_lin + 2.4)-(l_margem_col + 12, l_margem_lin + 4)
    'ImprimeString "Tanque", l_margem_col + 0.2, l_margem_lin + 2.65, l_local
    'ImprimeString "Tanque", l_margem_col + 2.2, l_margem_lin + 2.65, l_local
    'ImprimeString "Tanque", l_margem_col + 4.2, l_margem_lin + 2.65, l_local
    'ImprimeString "Tanque", l_margem_col + 6.2, l_margem_lin + 2.65, l_local
    'ImprimeString "Tanque", l_margem_col + 8.2, l_margem_lin + 2.65, l_local
    'ImprimeString "Tanque", l_margem_col + 10.2, l_margem_lin + 2.65, l_local
    'ImprimeString "3.1 Estoque de abertura", l_margem_col + 12.2, l_margem_lin + 2.65, l_local
    
    'Printer.Line (l_margem_col, l_margem_lin + 4)-(l_col_max, l_margem_lin + 4)
    
    'Printer.Line (l_margem_col, l_margem_lin + 4.8)-(l_col_max, l_margem_lin + 4.8)
    'Printer.Line (l_margem_col + 13.6, l_margem_lin + 4)-(l_margem_col + 13.6, l_margem_lin + 15.1)
    'ImprimeString "4 Volume recebido no dia (em litros)", l_margem_col + 0.1, l_margem_lin + 4.25, l_local
    'ImprimeString "4.1 Nr. tanque descarga", l_margem_col + 10.1, l_margem_lin + 4.25, l_local
    'ImprimeString "4.2 Volume recebido", l_margem_col + 13.7, l_margem_lin + 4.25, l_local
    
    'Printer.Line (l_margem_col, l_margem_lin + 6.5)-(l_col_max, l_margem_lin + 6.5)
    'ImprimeString "Nota fiscal nr.:", l_margem_col + 0.1, l_margem_lin + 4.95, l_local
    'ImprimeString "de", l_margem_col + 6.6, l_margem_lin + 4.95, l_local
    'ImprimeString "Nota fiscal nr.:", l_margem_col + 0.1, l_margem_lin + 5.45, l_local
    'ImprimeString "de", l_margem_col + 6.6, l_margem_lin + 5.45, l_local
    'ImprimeString "Nota fiscal nr.:", l_margem_col + 0.1, l_margem_lin + 5.95, l_local
    'ImprimeString "de", l_margem_col + 6.6, l_margem_lin + 5.95, l_local
    
    'Printer.Line (l_margem_col, l_margem_lin + 7.3)-(l_col_max, l_margem_lin + 7.3)
    'ImprimeString "4.3 Total recebido", l_margem_col + 10.1, l_margem_lin + 6.75, l_local
    
    'Printer.Line (l_margem_col, l_margem_lin + 8.5)-(l_col_max, l_margem_lin + 8.5)
    'ImprimeString "5 Volume vendido no dia (em litros)", l_margem_col + 0.1, l_margem_lin + 7.75, l_local
    'ImprimeString "4.4 Volume disponível", l_margem_col + 10.1, l_margem_lin + 7.45, l_local
    'ImprimeString "(3.1 + 4.3)", l_margem_col + 11, l_margem_lin + 8.05, l_local
    
    'Printer.Line (l_margem_col, l_margem_lin + 9.3)-(l_col_max, l_margem_lin + 9.3)
    'Printer.Line (l_margem_col + 2, l_margem_lin + 8.5)-(l_margem_col + 2, l_margem_lin + 12.3)
    'Printer.Line (l_margem_col + 4.5, l_margem_lin + 8.5)-(l_margem_col + 4.5, l_margem_lin + 12.3)
    'Printer.Line (l_margem_col + 7.2, l_margem_lin + 8.5)-(l_margem_col + 7.2, l_margem_lin + 12.3)
    'ImprimeString "5.1 Tanque", l_margem_col + 0.1, l_margem_lin + 8.75, l_local
    'ImprimeString "5.2 Bico", l_margem_col + 2.1, l_margem_lin + 8.75, l_local
    'ImprimeString "5.3 + Fechamento", l_margem_col + 4.6, l_margem_lin + 8.75, l_local
    'ImprimeString "5.4 - Abertura", l_margem_col + 7.3, l_margem_lin + 8.75, l_local
    'ImprimeString "5.5 - Aferição", l_margem_col + 10.1, l_margem_lin + 8.75, l_local
    'ImprimeString "5.6 = vendas no bico", l_margem_col + 13.7, l_margem_lin + 8.75, l_local
    
    'Printer.Line (l_margem_col, l_margem_lin + 12.3)-(l_col_max, l_margem_lin + 12.3)
    
    'Printer.Line (l_margem_col, l_margem_lin + 13.1)-(l_col_max, l_margem_lin + 13.1)
    'ImprimeString "10 Valor das vendas", l_margem_col + 0.1, l_margem_lin + 12.55, l_local
    'ImprimeString "5.7 Vendas no dia", l_margem_col + 10.1, l_margem_lin + 12.55, l_local
    
    'Printer.Line (l_margem_col, l_margem_lin + 14.3)-(l_col_max, l_margem_lin + 14.3)
    'ImprimeString "10.1 Valor das vendas do dia", l_margem_col + 0.1, l_margem_lin + 13.25, l_local
    'ImprimeString "(5.7 x Preço bomba)", l_margem_col + 1.1, l_margem_lin + 13.75, l_local
    'ImprimeString "6 Estoque escritural", l_margem_col + 10.1, l_margem_lin + 13.25, l_local
    'ImprimeString "(4.4 - 5.7)", l_margem_col + 11, l_margem_lin + 13.75, l_local
    
    'Printer.Line (l_margem_col, l_margem_lin + 15.1)-(l_col_max, l_margem_lin + 15.1)
    'ImprimeString "10.2 Valor acumulado no mês", l_margem_col + 0.1, l_margem_lin + 14.55, l_local
    'ImprimeString "7 Estoque fechamento (9.1)", l_margem_col + 10.1, l_margem_lin + 14.55, l_local
    
    'Printer.Line (l_margem_col + 10, l_margem_lin + 15.9)-(l_col_max, l_margem_lin + 15.9)
    'Printer.Line (l_margem_col, l_margem_lin + 16.4)-(l_margem_col + 10, l_margem_lin + 16.4)
    'ImprimeString "11 Para uso do revendedor", l_margem_col + 0.1, l_margem_lin + 15.35, l_local
    'ImprimeString "8 - Perdas + sobras (*)", l_margem_col + 10.1, l_margem_lin + 15.35, l_local
    'ImprimeString "12 Destinado a fiscalização DNC", l_margem_col + 10.1, l_margem_lin + 16.1, l_local
    
    'Printer.Line (l_margem_col + 10, l_margem_lin + 18.8)-(l_col_max, l_margem_lin + 18.8)
    'Printer.Line (l_margem_col, l_margem_lin + 21.9)-(l_col_max, l_margem_lin + 21.9)
    'ImprimeString "13 Observações", l_margem_col + 0.1, l_margem_lin + 16.6, l_local
    'ImprimeString "Outros orgãos fiscais", l_margem_col + 10.1, l_margem_lin + 19, l_local
    
    'Printer.Line (l_margem_col, l_margem_lin + 22.7)-(l_col_max, l_margem_lin + 22.7)
    'ImprimeString "Conciliação dos Estoques", l_margem_col + 7.3, l_margem_lin + 22.1, l_local
   
    'Printer.Line (l_margem_col + 3.4, l_margem_lin + 23.5)-(l_col_max, l_margem_lin + 23.5)
    'Printer.Line (l_margem_col, l_margem_lin + 24.3)-(l_col_max, l_margem_lin + 24.3)
    'Printer.Line (l_margem_col + 3.4, l_margem_lin + 22.7)-(l_margem_col + 3.4, l_margem_lin + 24.3)
    'Printer.Line (l_margem_col + 4.8, l_margem_lin + 22.7)-(l_margem_col + 4.8, l_margem_lin + 23.5)
    'Printer.Line (l_margem_col + 5.4, l_margem_lin + 22.7)-(l_margem_col + 5.4, l_margem_lin + 24.3)
    'Printer.Line (l_margem_col + 6.8, l_margem_lin + 22.7)-(l_margem_col + 6.8, l_margem_lin + 23.5)
    'Printer.Line (l_margem_col + 7.4, l_margem_lin + 22.7)-(l_margem_col + 7.4, l_margem_lin + 24.3)
    'Printer.Line (l_margem_col + 8.8, l_margem_lin + 22.7)-(l_margem_col + 8.8, l_margem_lin + 23.5)
    'Printer.Line (l_margem_col + 9.4, l_margem_lin + 22.7)-(l_margem_col + 9.4, l_margem_lin + 24.3)
    'Printer.Line (l_margem_col + 10.8, l_margem_lin + 22.7)-(l_margem_col + 10.8, l_margem_lin + 23.5)
    'Printer.Line (l_margem_col + 11.4, l_margem_lin + 22.7)-(l_margem_col + 11.4, l_margem_lin + 24.3)
    'Printer.Line (l_margem_col + 12.8, l_margem_lin + 22.7)-(l_margem_col + 12.8, l_margem_lin + 23.5)
    'Printer.Line (l_margem_col + 13.4, l_margem_lin + 22.7)-(l_margem_col + 13.4, l_margem_lin + 24.3)
    'Printer.Line (l_margem_col + 14.8, l_margem_lin + 22.7)-(l_margem_col + 14.8, l_margem_lin + 23.5)
    'Printer.Line (l_margem_col + 15.4, l_margem_lin + 22.7)-(l_margem_col + 15.4, l_margem_lin + 24.3)
    'ImprimeString "9 Fechamento físico", l_margem_col + 0.1, l_margem_lin + 23.2, l_local
    'ImprimeString "Tanque", l_margem_col + 3.6, l_margem_lin + 22.9, l_local
    'ImprimeString "Tanque", l_margem_col + 5.6, l_margem_lin + 22.9, l_local
    'ImprimeString "Tanque", l_margem_col + 7.6, l_margem_lin + 22.9, l_local
    'ImprimeString "Tanque", l_margem_col + 9.6, l_margem_lin + 22.9, l_local
    'ImprimeString "Tanque", l_margem_col + 11.6, l_margem_lin + 22.9, l_local
    'ImprimeString "Tanque", l_margem_col + 13.6, l_margem_lin + 22.9, l_local
    'ImprimeString "9.1  Total", l_margem_col + 15.6, l_margem_lin + 22.9, l_local
    'ImprimeString "(*) Atenção se o resultado for negativo, pode estar havendo vazamento do produto para o meio ambiente", l_margem_col + 0.5, l_margem_lin + 24.6, l_local
End Sub
Private Sub Relatorio()
    ZeraVariaveis
    'seleciona medidas para centímetros
    Printer.ScaleMode = 7
    'Seleciona Formulário de cheque
    'Printer.PaperSize = 9
    'Seleciona largura do formulário
    Printer.ScaleWidth = 210
    'l_lin_max = Printer.ScaleWidth
    'Seleciona altura do formulário
    Printer.ScaleHeight = 280
    Printer.ScaleMode = 7
    'l_lin_max = Printer.ScaleHeight
    'Seleciona nome da fonte
    Printer.FontName = "Arial"
    Printer.FontName = "Arial"
    ImprimeGrade
   'ImprimeDados
    Printer.EndDoc
    cmd_sair.SetFocus
End Sub
Private Sub cmd_imprimir_Click()
    If ValidaCampos Then
        If SelecionaImpressoraHP(Me) Then
            Relatorio
        End If
    End If
End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If Not Val(txt_pagina) > 0 Then
        MsgBox "Informe o número de páginas a serem impressas.", 64, "Atenção!"
        txt_pagina.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub Form_Load()
    Screen.MousePointer = 1
    CentraForm Me
End Sub
Private Sub txt_pagina_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_imprimir.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_pagina_LostFocus()
    txt_pagina = Format(Val(txt_pagina), "000")
End Sub
