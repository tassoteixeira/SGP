VERSION 5.00
Begin VB.Form frmAguarde 
   AutoRedraw      =   -1  'True
   Caption         =   "Aguarde! Processando..."
   ClientHeight    =   2400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4575
   ControlBox      =   0   'False
   Icon            =   "frmAguarde.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   1680
      Top             =   960
   End
   Begin VB.Label lblContador 
      Alignment       =   2  'Center
      Caption         =   "lblContador"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   615
      Left            =   2880
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblMensagem 
      Alignment       =   2  'Center
      Caption         =   "lblMensagem"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   1380
      Width           =   4035
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      Caption         =   "lblTitulo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   420
      Width           =   3975
   End
End
Attribute VB_Name = "frmAguarde"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lNumero As Integer
Private Aliquota As New cAliquota

Public Sub Finaliza()
    'Call TestaConexao(83, "frmAguarde.Finaliza 1")
    Timer1.Enabled = False
    Timer1.Interval = 0
    'Call TestaConexao(84, "frmAguarde.Finaliza 2")
    Set Aliquota = Nothing
    Unload Me
End Sub
Public Sub MostraMensagens(ByVal pTitulo As String, ByVal pMensagem As String, ByVal pSuperior As Currency, ByVal pEsquerda As Currency, ByVal pLargura As Currency, ByVal pAltura As Currency)
On Error GoTo trata_erro
    'Call TestaConexao(81, "frmAguarde.MostraMensagens 0 - pTitulo=" & pTitulo)
    'Call TestaConexao(81, "frmAguarde.MostraMensagens 0 - pMensagem=" & pMensagem)
    'Call TestaConexao(81, "frmAguarde.MostraMensagens 0 - pSuperior=" & pSuperior)
    'Call TestaConexao(81, "frmAguarde.MostraMensagens 0 - pEsquerda=" & pEsquerda)
    'Call TestaConexao(81, "frmAguarde.MostraMensagens 0 - pLargura=" & pLargura)
    'Call TestaConexao(81, "frmAguarde.MostraMensagens 0 - pAltura=" & pAltura)
    'Call TestaConexao(81, "frmAguarde.MostraMensagens 1 - Timer1.Enabled=" & Timer1.Enabled)
    Me.Top = pSuperior
    Me.Left = pEsquerda
    Me.Width = pLargura
    Me.Height = pAltura
    'Call TestaConexao(81, "frmAguarde.MostraMensagens 2 - Timer1.Enabled=" & Timer1.Enabled)
    DoEvents
    'Call TestaConexao(81, "frmAguarde.MostraMensagens 3 - Timer1.Enabled=" & Timer1.Enabled)
    
    'CentraForm Me
    lblTitulo.Width = pLargura
    lblTitulo.Top = pAltura * 0.25
    lblMensagem.Width = pLargura
    lblMensagem.Top = pAltura * 0.6
    'Call TestaConexao(81, "frmAguarde.MostraMensagens 4 - Timer1.Enabled=" & Timer1.Enabled)
    DoEvents
    'Call TestaConexao(81, "frmAguarde.MostraMensagens 5 - Timer1.Enabled=" & Timer1.Enabled)
    lblTitulo.Caption = pTitulo
    lblMensagem.Caption = pMensagem
    'Call TestaConexao(81, "frmAguarde.MostraMensagens 6 - Timer1.Enabled=" & Timer1.Enabled)
    DoEvents
    'Call TestaConexao(81, "frmAguarde.MostraMensagens 7 - Timer1.Enabled=" & Timer1.Enabled)
    Exit Sub

trata_erro:
    Call CriaLogCupom("Erro frmAguarde.MostraMensagens: Erro=" & Err.Number & " - " & Err.Description)
End Sub
Public Sub IniciaContador(ByVal pNumeroInicial As Integer, Optional pEspacoContador As Integer = 1000)
    'Call TestaConexao(82, "frmAguarde.IniciaContador 1")
    lblContador.Visible = True
    lblContador.Width = Me.Width
    lblContador.Top = lblMensagem.Top + pEspacoContador
    lblContador.Left = 200
    lNumero = pNumeroInicial
    lblContador.Caption = lNumero
    DoEvents
    Timer1.Enabled = True
    Timer1.Interval = 1000
    'Call TestaConexao(82, "frmAguarde.IniciaContador 2")
End Sub
Private Sub Form_Load()
    'Call TestaConexao(80, "frmAguarde.Form_Load 1 - Timer1.Enabled=" & Timer1.Enabled)
    lblTitulo.Caption = ""
    lblMensagem.Caption = ""
    lblContador.Caption = ""
    Timer1.Enabled = False
    Timer1.Interval = 0
    'Call TestaConexao(80, "frmAguarde.Form_Load 2 - Timer1.Enabled=" & Timer1.Enabled)
End Sub
Private Sub Timer1_Timer()
On Error GoTo trata_erro
    lNumero = lNumero - 1
    lblContador.Caption = lNumero
    DoEvents
    Exit Sub

trata_erro:
    Call CriaLogCupom("Erro Timer1_Timer: Erro=" & Err.Number & " - " & Err.Description)
End Sub
Private Function TestaConexao(ByVal pNumeroTeste As Integer, ByVal pVindoDe As String) As Boolean
    Call CriaLogCupom("TestaConexao - Iniciado. pNumeroTeste=" & pNumeroTeste & " - Vindo de: " & pVindoDe)
    TestaConexao = False
    If Aliquota.LocalizarUltimo() Then
        TestaConexao = True
        Call CriaLogCupom("TestaConexao - CONEXÃO OK. pNumeroTeste=" & pNumeroTeste)
    Else
        Call CriaLogCupom("TestaConexao - ERRO DE CONEXÃO. pNumeroTeste=" & pNumeroTeste)
    End If
    Call CriaLogCupom("TestaConexao - Finalizado. pNumeroTeste=" & pNumeroTeste)
End Function

